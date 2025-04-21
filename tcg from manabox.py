import argparse
import pandas as pd


# -----------------------------------------
# Script: map_quantities.py
# Description:
#   Reads an Excel workbook with two sheets (primary and secondary),
#   maps "Quantity" from the secondary sheet onto the primary sheet
#   based on matching Number/Collector number and foil status,
#   and writes the result back to a new Excel file.
# -----------------------------------------

def parse_args():
    parser = argparse.ArgumentParser(
        description='Map quantities from inventory sheet to listings sheet.'
    )
    parser.add_argument(
        '-i', '--input',
        required=True,
        help='Path to the input Excel file'
    )
    parser.add_argument(
        '-o', '--output',
        required=True,
        help='Path for the output Excel file'
    )
    parser.add_argument(
        '--primary-sheet',
        default='Foundations 20250421_100121',
        help='Name of the primary sheet (listings)'
    )
    parser.add_argument(
        '--secondary-sheet',
        default='Manabox scanned',
        help='Name of the secondary sheet (inventory)'
    )
    parser.add_argument(
        '--number-col',
        default='Number',
        help='Column name in primary sheet for card number'
    )
    parser.add_argument(
        '--collector-col',
        default='Collector number',
        help='Column name in secondary sheet for collector number'
    )
    parser.add_argument(
        '--condition-col',
        default='Condition',
        help='Column name in primary sheet for condition'
    )
    parser.add_argument(
        '--foil-keyword',
        default='foil',
        help='Keyword to detect foil in the condition field'
    )
    parser.add_argument(
        '--quantity-col',
        default='Quantity',
        help='Column name in secondary sheet for inventory quantity'
    )
    parser.add_argument(
        '--output-sheet',
        default='Foundations Updated',
        help='Name of the sheet to write the updated primary data'
    )
    return parser.parse_args()


def main():
    args = parse_args()

    # Load sheets
    xls = pd.ExcelFile(args.input)
    df_primary = xls.parse(args.primary_sheet)
    df_secondary = xls.parse(args.secondary_sheet)

    # Prepare join keys
    df_primary['__num__'] = df_primary[args.number_col].astype(str).str.strip()
    df_secondary['__collector__'] = df_secondary[args.collector_col].astype(str).str.strip()

    # Determine foil flag: 'foil' if condition contains the keyword, else 'normal'
    df_primary['__foil_flag__'] = (
        df_primary[args.condition_col]
        .str.lower()
        .str.contains(args.foil_keyword.lower())
        .map({True: 'foil', False: 'normal'})
    )

    # Build mapping (collector, foil) -> quantity
    key_cols = ['__collector__', 'Foil']
    mapping = (
        df_secondary
        .set_index(key_cols)[args.quantity_col]
        .to_dict()
    )

    # Apply mapping to get Add to Quantity
    df_primary['Add to Quantity'] = df_primary.apply(
        lambda row: mapping.get((row['__num__'], row['__foil_flag__']), 0),
        axis=1
    )

    # Clean up helper columns
    df_primary.drop(columns=['__num__', '__foil_flag__'], inplace=True)

    # Write results
    with pd.ExcelWriter(args.output, engine='openpyxl') as writer:
        df_primary.to_excel(writer, sheet_name=args.output_sheet, index=False)
        df_secondary.to_excel(writer, sheet_name=args.secondary_sheet, index=False)

    print(f"Updated workbook written to: {args.output}")


if __name__ == '__main__':
    main()
