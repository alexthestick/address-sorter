#!/usr/bin/env python3
"""
Address Sorter - Automates sorting of addresses into different categories
with ROE deduplication logic and anomaly detection.
"""

import pandas as pd
import numpy as np
from collections import defaultdict, Counter
import re
import sys
import os


class AddressSorter:
    def __init__(self, input_file):
        """Initialize the address sorter with input file."""
        self.input_file = input_file
        self.df = None
        self.tabs = {
            'All': None,
            'Public': None,
            'Commercial': None,
            'ROE': None,
            'Competitive': None,
            'Other': None,
            'Remove': None,
            'Flagged for Review': None,
            'Unit Count': None
        }
        self.flagged_addresses = []
        self.required_columns = ['ID', 'Street Address', 'Unit Number', 'Building Type', 'Subname']

    def load_data(self):
        """Load the input CSV file."""
        print(f"Loading data from {self.input_file}...")
        if self.input_file.endswith('.csv'):
            self.df = pd.read_csv(self.input_file, low_memory=False)
        elif self.input_file.endswith('.xlsx'):
            self.df = pd.read_excel(self.input_file)
        else:
            raise ValueError("Input file must be CSV or Excel (.xlsx)")

        # Check for required columns
        missing_cols = [col for col in self.required_columns if col not in self.df.columns]
        if missing_cols:
            raise ValueError(f"Missing required columns: {missing_cols}")

        # Keep only essential columns (plus a few useful ones)
        essential_cols = self.required_columns.copy()
        optional_cols = ['City', 'Zip', 'Plus 4 Code', 'Zone', 'Street Name']
        for col in optional_cols:
            if col in self.df.columns:
                essential_cols.append(col)

        self.df = self.df[essential_cols].copy()

        # Handle no-subname: assign placeholder
        self.df['Subname'] = self.df['Subname'].fillna('No Subname')
        self.df['Subname'] = self.df['Subname'].replace('', 'No Subname')

        print(f"Loaded {len(self.df)} addresses")
        print(f"Building types found: {self.df['Building Type'].value_counts().to_dict()}")

    def initial_sort(self):
        """Perform initial sorting into basic categories."""
        print("\nPerforming initial sort by building type...")

        # All tab gets everything
        self.tabs['All'] = self.df.copy()

        # Public = Residential only (not MDU, not SFA, not HOA, not Mobile)
        self.tabs['Public'] = self.df[self.df['Building Type'] == 'Residential'].copy()

        # Commercial
        self.tabs['Commercial'] = self.df[self.df['Building Type'] == 'Commercial'].copy()

        # Competitive
        self.tabs['Competitive'] = self.df[self.df['Building Type'] == 'Competitive'].copy()

        # Other
        self.tabs['Other'] = self.df[self.df['Building Type'] == 'Other'].copy()

        # ROE candidates (before deduplication)
        roe_types = ['Residential - MDU', 'SFA', 'HOA', 'Mobile']
        roe_candidates = self.df[self.df['Building Type'].isin(roe_types)].copy()

        print(f"  Public: {len(self.tabs['Public'])} addresses")
        print(f"  Commercial: {len(self.tabs['Commercial'])} addresses")
        print(f"  Competitive: {len(self.tabs['Competitive'])} addresses")
        print(f"  Other: {len(self.tabs['Other'])} addresses")
        print(f"  ROE candidates (before deduplication): {len(roe_candidates)} addresses")

        return roe_candidates

    def detect_unit_anomalies(self, unit_value):
        """Detect anomalies in unit numbers."""
        if pd.isna(unit_value):
            return None

        unit_str = str(unit_value).upper().strip()

        # Check for office/commercial indicators
        if any(keyword in unit_str for keyword in ['OFC', 'OFFICE', 'CLUBHOUSE', 'LEASING', 'CLUB']):
            return 'office_indicator'

        # Check for suite (commercial)
        if 'STE' in unit_str or 'SUITE' in unit_str:
            return 'suite_indicator'

        return None

    def normalize_unit_format(self, unit_value):
        """Normalize unit format for comparison."""
        if pd.isna(unit_value):
            return None

        unit_str = str(unit_value).strip()

        # Extract just the number if possible
        numbers = re.findall(r'\d+', unit_str)
        if numbers:
            return numbers[0]

        return unit_str

    def get_unit_format_type(self, unit_value):
        """Categorize the unit format type."""
        if pd.isna(unit_value):
            return 'no_unit'

        unit_str = str(unit_value).strip().upper()

        # Standard formats
        if unit_str.startswith('UNIT'):
            return 'unit_format'  # e.g., "UNIT 6"
        elif unit_str.startswith('APT'):
            return 'apt_format'   # e.g., "APT 6"
        elif re.match(r'^\d+$', unit_str):
            return 'number_only'  # e.g., "6"

        # Anomalous formats that should be flagged
        elif unit_str.startswith('STE'):
            return 'ste_format'   # e.g., "STE 100" (commercial)
        elif 'BLDG' in unit_str or 'BUILDING' in unit_str:
            return 'building_format'  # e.g., "BLDG J"
        elif unit_str.startswith('#') and 'B-' in unit_str:
            return 'hash_b_format'  # e.g., "# B-1234"
        elif unit_str.startswith('#'):
            return 'hash_format'  # e.g., "# 6"
        else:
            return 'other_format'

    def process_roe_subname(self, subname_df, subname, building_type):
        """Process a single subname group for ROE logic."""
        keep_addresses = []
        remove_addresses = []
        flagged = []

        # Check for unit anomalies first (OFC, OFFICE, STE, etc.)
        anomaly_indices = []
        for idx, row in subname_df.iterrows():
            anomaly = self.detect_unit_anomalies(row['Unit Number'])
            if anomaly:
                anomaly_indices.append(idx)
                continue

        # Remove anomalies from consideration (they go straight to Remove, not flagged)
        subname_df = subname_df.drop(anomaly_indices)
        remove_addresses.extend(anomaly_indices)  # Add to remove list

        if len(subname_df) == 0:
            return keep_addresses, remove_addresses, flagged

        # Check for 5-digit Plus 4 Code in ROE addresses (dud addresses like backyards)
        # Apply to HOA, MDU, and SFA building types
        plus4_anomaly_indices = []
        roe_types = ['HOA', 'Residential - MDU', 'SFA']
        if building_type in roe_types and 'Plus 4 Code' in subname_df.columns:
            for idx, row in subname_df.iterrows():
                plus4 = row['Plus 4 Code']
                zip_code = row.get('Zip', None)
                if pd.notna(plus4):
                    plus4_str = str(plus4).strip()
                    zip_str = str(zip_code).strip() if pd.notna(zip_code) else ''
                    # Check if Plus 4 Code is 5 digits (should be 4) or equals Zip
                    if (len(plus4_str) == 5 and plus4_str.isdigit()) or (plus4_str == zip_str):
                        plus4_anomaly_indices.append(idx)

        # Remove Plus 4 Code anomalies
        subname_df = subname_df.drop(plus4_anomaly_indices)
        remove_addresses.extend(plus4_anomaly_indices)

        if len(subname_df) == 0:
            return keep_addresses, remove_addresses, flagged

        # Check for oversized communities (>800 addresses)
        if len(subname_df) > 800:
            for idx, row in subname_df.iterrows():
                flagged.append({
                    'index': idx,
                    'reason': f'Community has {len(subname_df)} addresses (>800 threshold)',
                    **row.to_dict()
                })

        # DIFFERENT LOGIC FOR MDU vs SFA/HOA
        is_mdu = building_type == 'Residential - MDU'

        # SPECIAL CASE: Check if this is a "condo-style" SFA/HOA
        # (all addresses share the same street address with different units)
        unique_streets = subname_df['Street Address'].nunique()
        total_addresses = len(subname_df)
        with_units = len(subname_df[~subname_df['Unit Number'].isna()])

        # If 80%+ addresses share the same street AND have units, treat like MDU
        is_condo_style = (unique_streets <= 3 and
                         with_units / total_addresses > 0.8 and
                         total_addresses > 50)

        if is_condo_style and not is_mdu:
            print(f"    Note: Detected condo-style {building_type} (treating like MDU)")

        # FOR MDU: Check for mixed unit formats BEFORE deduplication
        # Analyze unit format consistency within subname
        all_formats = subname_df['Unit Number'].apply(self.get_unit_format_type)
        format_counts = Counter(all_formats)

        # Define standard vs anomalous formats
        standard_formats = {'unit_format', 'apt_format', 'number_only'}
        anomalous_formats = {'ste_format', 'building_format', 'hash_b_format', 'hash_format', 'other_format'}

        # Track which indices have non-majority formats (for MDU only)
        non_majority_format_indices = []
        if is_mdu and len(format_counts) > 1:
            # Count standard formats only (exclude no_unit)
            standard_format_counts = {fmt: count for fmt, count in format_counts.items()
                                     if fmt in standard_formats}

            # If we have standard formats, find the most common one
            if standard_format_counts:
                majority_standard_format = max(standard_format_counts, key=standard_format_counts.get)

                # Flag ALL addresses that don't match the majority standard format
                # This includes both anomalous formats AND minority standard formats
                for idx, row in subname_df.iterrows():
                    row_format = self.get_unit_format_type(row['Unit Number'])
                    if row_format != majority_standard_format and row_format != 'no_unit':
                        non_majority_format_indices.append(idx)

                        # Determine if it's anomalous or just minority standard
                        if row_format in anomalous_formats:
                            reason = f'MDU anomalous format: {row_format} (standard is {majority_standard_format})'
                        else:
                            reason = f'MDU minority format: {row_format} (majority is {majority_standard_format})'

                        flagged.append({
                            'index': idx,
                            'reason': reason,
                            **row.to_dict()
                        })

        # Remove non-majority format addresses from consideration in deduplication
        subname_df = subname_df.drop(non_majority_format_indices)

        if len(subname_df) == 0:
            return keep_addresses, remove_addresses, flagged

        # Group by street address to find duplicates
        address_groups = subname_df.groupby('Street Address')

        for street_address, group in address_groups:
            group_indices = group.index.tolist()

            # Separate those with unit numbers and those without
            no_unit = group[group['Unit Number'].isna()]
            with_unit = group[~group['Unit Number'].isna()]

            no_unit_indices = no_unit.index.tolist()
            with_unit_indices = with_unit.index.tolist()

            if is_mdu or is_condo_style:
                # MDU LOGIC (or condo-style SFA): Each unit is a separate customer
                # If we have both versions (with and without unit)
                if len(no_unit_indices) > 0 and len(with_unit_indices) > 0:
                    # Keep all with-unit versions (actual apartments/condos)
                    keep_addresses.extend(with_unit_indices)
                    # Remove no-unit versions (likely leasing office/clubhouse)
                    remove_addresses.extend(no_unit_indices)

                # If only with-unit versions exist
                elif len(with_unit_indices) > 0 and len(no_unit_indices) == 0:
                    # Keep all unit versions
                    keep_addresses.extend(with_unit_indices)

                # If only no-unit versions exist
                elif len(no_unit_indices) > 0 and len(with_unit_indices) == 0:
                    # Check if this is an isolated no-unit in a mostly with-unit community
                    # Calculate percentage of addresses with units in the whole subname
                    total_with_units = len(subname_df[~subname_df['Unit Number'].isna()])
                    total_in_subname = len(subname_df)
                    percent_with_units = total_with_units / total_in_subname if total_in_subname > 0 else 0

                    # If 80%+ of community has units, remove isolated no-unit addresses (likely offices)
                    if percent_with_units >= 0.8:
                        remove_addresses.extend(no_unit_indices)
                    else:
                        # Otherwise keep them (might be valid addresses like Montelena)
                        keep_addresses.extend(no_unit_indices)

            else:
                # SFA/HOA LOGIC: Prefer addresses WITHOUT units, but keep unique WITH units
                # If we have both versions (with and without unit)
                if len(no_unit_indices) > 0 and len(with_unit_indices) > 0:
                    # Keep all no-unit versions
                    keep_addresses.extend(no_unit_indices)
                    # Remove all with-unit versions (duplicates)
                    remove_addresses.extend(with_unit_indices)

                # If only no-unit versions exist
                elif len(no_unit_indices) > 0 and len(with_unit_indices) == 0:
                    keep_addresses.extend(no_unit_indices)

                # If only with-unit versions exist
                elif len(no_unit_indices) == 0 and len(with_unit_indices) > 0:
                    # Keep all unique unit versions
                    keep_addresses.extend(with_unit_indices)

        # For non-MDU: Check if there's a clear majority format and minority formats
        # (MDU format checking already happened earlier)
        if not is_mdu and len(format_counts) > 1:
            total = len(all_formats)
            majority_format = format_counts.most_common(1)[0][0]
            majority_count = format_counts[majority_format]

            # If majority is > 80%, flag minority formats
            if majority_count / total > 0.8:
                for idx, row in subname_df.iterrows():
                    if idx not in keep_addresses and idx not in remove_addresses:
                        row_format = self.get_unit_format_type(row['Unit Number'])
                        if row_format != majority_format and row_format != 'no_unit':
                            flagged.append({
                                'index': idx,
                                'reason': f'Minority unit format ({row_format}) vs majority ({majority_format})',
                                **row.to_dict()
                            })

        # Check for one-off scenarios
        total_in_subname = len(subname_df)
        no_unit_count = len(subname_df[subname_df['Unit Number'].isna()])
        with_unit_count = total_in_subname - no_unit_count

        # Scenario: 149 with units, 1 without (or vice versa)
        if total_in_subname > 10:  # Only flag if substantial sample size
            if no_unit_count == 1 and with_unit_count > 10:
                # The single no-unit is likely office/clubhouse
                single_idx = subname_df[subname_df['Unit Number'].isna()].index[0]
                if single_idx in keep_addresses:
                    keep_addresses.remove(single_idx)
                remove_addresses.append(single_idx)
                flagged.append({
                    'index': single_idx,
                    'reason': f'One-off: Single address without unit among {with_unit_count} with units',
                    **subname_df.loc[single_idx].to_dict()
                })
            elif with_unit_count == 1 and no_unit_count > 10:
                # The single with-unit is likely office
                single_idx = subname_df[~subname_df['Unit Number'].isna()].index[0]
                if single_idx in keep_addresses:
                    keep_addresses.remove(single_idx)
                remove_addresses.append(single_idx)
                flagged.append({
                    'index': single_idx,
                    'reason': f'One-off: Single address with unit among {no_unit_count} without units',
                    **subname_df.loc[single_idx].to_dict()
                })

        return keep_addresses, remove_addresses, flagged

    def process_roe_deduplication(self, roe_candidates):
        """Process ROE candidates with deduplication logic."""
        print("\nProcessing ROE deduplication logic...")

        keep_indices = []
        remove_indices = []

        # Group by Subname AND Building Type (to handle "No Subname" properly)
        grouped = roe_candidates.groupby(['Subname', 'Building Type'])

        for (subname, building_type), subname_df in grouped:
            print(f"\n  Processing: {subname} ({building_type}) - {len(subname_df)} addresses")

            # For "No Subname", sort by Street Name to group related addresses together
            if subname == 'No Subname':
                if 'Street Name' in subname_df.columns:
                    subname_df = subname_df.sort_values('Street Name')
                else:
                    # Extract street name from Street Address if column doesn't exist
                    subname_df = subname_df.sort_values('Street Address')

            keep, remove, flagged = self.process_roe_subname(subname_df, subname, building_type)

            keep_indices.extend(keep)
            remove_indices.extend(remove)
            self.flagged_addresses.extend(flagged)

            print(f"    Keeping: {len(keep)} addresses")
            print(f"    Removing: {len(remove)} addresses")
            print(f"    Flagged: {len(flagged)} addresses")

        # Create ROE and Remove tabs
        self.tabs['ROE'] = roe_candidates.loc[keep_indices].copy()
        self.tabs['Remove'] = roe_candidates.loc[remove_indices].copy()

        # Sort ROE by Subname, then Street Address (or Street Name)
        sort_cols = ['Subname']
        if 'Street Name' in self.tabs['ROE'].columns:
            sort_cols.append('Street Name')
        sort_cols.append('Street Address')
        self.tabs['ROE'] = self.tabs['ROE'].sort_values(sort_cols)

        # Add blank rows between communities for better readability
        self.add_spacing_to_roe()

        print(f"\nFinal ROE count: {len(self.tabs['ROE'])} addresses")
        print(f"Final Remove count: {len(self.tabs['Remove'])} addresses")

    def add_spacing_to_roe(self):
        """Add blank rows between different communities in ROE tab for readability."""
        if self.tabs['ROE'] is None or len(self.tabs['ROE']) == 0:
            return

        # Add Unit Count column if it doesn't exist
        if 'Unit Count' not in self.tabs['ROE'].columns:
            self.tabs['ROE'].insert(0, 'Unit Count', None)

        # Calculate unit counts per subname
        subname_counts = self.tabs['ROE']['Subname'].value_counts().to_dict()

        # Create a new dataframe with spacing and unit counts
        rows_with_spacing = []
        prev_subname = None
        first_in_community = True

        for idx, row in self.tabs['ROE'].iterrows():
            current_subname = row['Subname']

            # Add blank row when subname changes (except for first entry)
            if prev_subname is not None and current_subname != prev_subname:
                # Create a blank row with NaN values
                blank_row = {col: None for col in self.tabs['ROE'].columns}
                rows_with_spacing.append(blank_row)
                first_in_community = True

            # Add unit count only on first row of each community
            row_dict = row.to_dict()
            if first_in_community and current_subname in subname_counts:
                row_dict['Unit Count'] = subname_counts[current_subname]
                first_in_community = False
            else:
                row_dict['Unit Count'] = None

            rows_with_spacing.append(row_dict)
            prev_subname = current_subname

        # Replace ROE tab with spaced version
        self.tabs['ROE'] = pd.DataFrame(rows_with_spacing)

    def create_flagged_tab(self):
        """Create the Flagged for Review tab."""
        if self.flagged_addresses:
            self.tabs['Flagged for Review'] = pd.DataFrame(self.flagged_addresses)
            print(f"\nCreated Flagged for Review tab with {len(self.flagged_addresses)} addresses")
        else:
            # Create empty dataframe with same columns
            self.tabs['Flagged for Review'] = pd.DataFrame(columns=list(self.df.columns) + ['reason'])
            print("\nNo addresses flagged for review")

    def create_unit_count_tab(self):
        """Create the Unit Count summary tab."""
        print("\nGenerating Unit Count summary...")

        counts = {
            'Category': [],
            'Count': []
        }

        # Total
        counts['Category'].append('Total')
        counts['Count'].append(len(self.tabs['All']))

        # Public
        counts['Category'].append('Public')
        counts['Count'].append(len(self.tabs['Public']))

        # Commercial
        counts['Category'].append('Commercial')
        counts['Count'].append(len(self.tabs['Commercial']))

        # ROE by type
        roe_by_type = self.tabs['ROE'].groupby('Building Type').size()
        for building_type, count in roe_by_type.items():
            counts['Category'].append(f'ROE - {building_type}')
            counts['Count'].append(count)

        # Total ROE
        counts['Category'].append('ROE - Total')
        counts['Count'].append(len(self.tabs['ROE']))

        # Competitive
        counts['Category'].append('Competitive')
        counts['Count'].append(len(self.tabs['Competitive']))

        # Other
        counts['Category'].append('Other')
        counts['Count'].append(len(self.tabs['Other']))

        # Remove
        counts['Category'].append('Remove')
        counts['Count'].append(len(self.tabs['Remove']))

        self.tabs['Unit Count'] = pd.DataFrame(counts)

        print("\nUnit Count Summary:")
        for cat, cnt in zip(counts['Category'], counts['Count']):
            print(f"  {cat}: {cnt}")

    def save_output(self, output_file):
        """Save all tabs to an Excel file."""
        print(f"\nSaving output to {output_file}...")

        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            for tab_name, df in self.tabs.items():
                if df is not None and not df.empty:
                    df.to_excel(writer, sheet_name=tab_name, index=False)
                    print(f"  Saved {tab_name} tab ({len(df)} rows)")

        print(f"\n✓ Successfully saved to {output_file}")

    def run(self, output_file):
        """Run the complete address sorting process."""
        print("="*60)
        print("ADDRESS SORTER")
        print("="*60)

        self.load_data()
        roe_candidates = self.initial_sort()
        self.process_roe_deduplication(roe_candidates)
        self.create_flagged_tab()
        self.create_unit_count_tab()
        self.save_output(output_file)

        print("\n" + "="*60)
        print("PROCESSING COMPLETE")
        print("="*60)


def main():
    if len(sys.argv) < 2:
        # No CLI args: try to open a simple GUI file picker (macOS-friendly)
        try:
            # Lazy import so headless/CLI environments aren't impacted
            import tkinter as tk
            from tkinter import filedialog

            root = tk.Tk()
            root.withdraw()

            print("No input file provided. Opening file picker...")
            input_file = filedialog.askopenfilename(
                title='Select input file (CSV or Excel)',
                filetypes=[('CSV files', '*.csv'), ('Excel files', '*.xlsx')]
            )

            if not input_file:
                print("No file selected. Exiting.")
                sys.exit(1)

            default_output = os.path.splitext(input_file)[0] + '_sorted.xlsx'

            output_file = filedialog.asksaveasfilename(
                title='Save output Excel as',
                defaultextension='.xlsx',
                initialfile=os.path.basename(default_output),
                filetypes=[('Excel files', '*.xlsx')]
            )

            if not output_file:
                output_file = default_output

        except Exception:
            # Fallback to CLI usage message if Tkinter isn't available
            print("Usage: python address_sorter.py <input_file> [output_file]")
            print("\nExample:")
            print("  python address_sorter.py input.csv output.xlsx")
            print("  python address_sorter.py input.xlsx")
            sys.exit(1)
    else:
        input_file = sys.argv[1]

        # Generate output filename if not provided
        if len(sys.argv) >= 3:
            output_file = sys.argv[2]
        else:
            output_file = input_file.rsplit('.', 1)[0] + '_sorted.xlsx'

    sorter = AddressSorter(input_file)
    sorter.run(output_file)


if __name__ == "__main__":
    main()
