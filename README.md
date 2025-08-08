# Atom Reorder Tool

A bash script tool for reordering atoms in Excel files according to specific cyclic patterns. This tool identifies H atoms that have been moved to the end of the file and redistributes them back into their correct positions following a defined molecular pattern.

## Features

- **Automatic atom type recognition**: Identifies O, C, and H atoms in Excel format
- **Cyclic pattern reordering**: Rearranges atoms according to specific molecular cycles
- **Multiple output modes**: Preview, verbose, and custom output file options
- **Error handling**: Comprehensive validation and dependency checking

## Cyclic Patterns

The tool implements three distinct cyclic patterns:

1. **First cycle**: `O H C H H C H H` (8 atoms)
2. **Middle cycles**: `O C H H C H H` (7 atoms each)
3. **Last cycle**: `O C H H C H H O H` (9 atoms)

## Installation

### Prerequisites

- Python 3.x
- Required Python packages:

```bash
pip3 install pandas openpyxl
```

### Setup

1. Clone the repository:

```bash
git clone https://github.com/ErdunE/atom-reorder-tool.git
cd atom-reorder-tool
```

2. Make the script executable:

```bash
chmod +x reorder_atoms.sh
```

## Usage

### Basic Usage

```bash
./reorder_atoms.sh input.xlsx
```

### Command Options

```bash
# Preview mode (no file output)
./reorder_atoms.sh -p input.xlsx

# Verbose output with detailed processing information
./reorder_atoms.sh -v input.xlsx

# Custom output filename
./reorder_atoms.sh -o custom_output.xlsx input.xlsx

# Combination of options
./reorder_atoms.sh -v -o result.xlsx input.xlsx
```

### Help

```bash
./reorder_atoms.sh -h
```

## Input Format

The tool expects Excel files (.xlsx/.xls) with the following format:

- **Column 1**: Atom names in format `ELEMENT(NUMBER)` (e.g., `O(1)`, `C(2)`, `H(3)`)
- **Column 2**: X coordinate
- **Column 3**: Y coordinate  
- **Column 4**: Z coordinate

## Output

The script generates a reordered Excel file where:

- H atoms are redistributed from the end back into the molecular structure
- Atoms follow the specified cyclic patterns
- Original coordinates are preserved
- File format remains consistent with input

## Example

### Input Structure

```
O(1)    770.925    8.818     132.235
C(2)    770.692    8.652     133.619
...
H(752)  770.611    9.675     131.935
H(753)  769.824    9.296     133.884
```

### Output Structure

```
O(1)    770.925    8.818     132.235
H(752)  770.611    9.675     131.935
C(2)    770.692    8.652     133.619
H(753)  769.824    9.296     133.884
...
```

## Files in Repository

- `reorder_atoms.sh` - Main bash script
- `input.xlsx` - Sample input file for testing
- `input_reordered.xlsx` - Example output file

## Algorithm Details

1. **Parsing**: Reads Excel file and separates atoms by type (O, C, H)
2. **Validation**: Ensures H atoms are at the end of the file
3. **Cycle Calculation**: Determines number of cycles needed based on total atom count
4. **Reordering**: Redistributes atoms according to the cyclic patterns:
   - First cycle: 8 atoms
   - Middle cycles: 248 cycles × 7 atoms each
   - Last cycle: 9 atoms
5. **Output**: Generates new Excel file with reordered structure

## Error Handling

The script includes comprehensive error checking for:

- Missing dependencies (Python, pandas, openpyxl)
- Invalid input files
- Incorrect atom formats
- Insufficient atom counts for pattern completion

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is open source and available under the [MIT License](LICENSE).

## Author

Created for molecular data processing and atom reorder tasks.
