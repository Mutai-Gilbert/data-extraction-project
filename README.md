
# Data Extraction from Word Document to Excel

This project is a Python-based solution that extracts data from a Microsoft Word document (`.docx`) and structures the data into an Excel file (`.xlsx`). The extracted data is stored in an organized format to facilitate further analysis or processing.

## Project Structure

```
data-extraction/
├── data/
│   ├── input/
│   │   └── sample.docx         # Word document containing the data to be extracted
│   ├── output/
│       └── structured_data.xlsx # Output Excel file with the extracted data
├── scripts/
│   └── extract_data.py         # Python script for data extraction
├── venv/                       # Virtual environment for project dependencies
└── README.md                   # Project documentation
```

## Features

- Extracts text content from paragraphs in a Word document.
- Saves extracted data into an organized Excel file.
- Automatically creates necessary output directories if they do not exist.

## Prerequisites

Make sure you have the following installed on your system:

- Python 3.8+
- [pip](https://pip.pypa.io/en/stable/installation/) (Python package manager)
- A virtual environment (optional but recommended)

## Installation

1. **Clone the repository:**
   ```bash
   git clone https://github.com/Mutai-Gilbert/data-extraction-project.git
   cd hatchet
   ```

2. **Set up a virtual environment:**
   ```bash
   python -m venv venv
   source venv/bin/activate   # On Windows use: venv\Scripts\activate
   ```

3. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

## Usage

1. **Prepare the Word document:**
   - Place your input Word document in the `data/input` directory.
   - Ensure the file is named `sample.docx` (or modify the script accordingly).

2. **Run the extraction script:**
   ```bash
   python scripts/extract_data.py
   ```

3. **Check the output:**
   - The extracted data will be saved as an Excel file at `data/output/structured_data.xlsx`.

## Configuration

- To modify the input file path, update the `input_file` variable in the `extract_data.py` script.
- To change the output file path, update the `output_file` variable in the same script.

## Dependencies

This project uses the following libraries:
- [python-docx](https://python-docx.readthedocs.io/en/latest/) for reading Word documents
- [pandas](https://pandas.pydata.org/) for data manipulation and exporting to Excel
- [openpyxl](https://openpyxl.readthedocs.io/en/stable/) for working with Excel files

You can install these dependencies with:
```bash
pip install python-docx pandas openpyxl
```

## Troubleshooting

- **ModuleNotFoundError:** Ensure all dependencies are installed correctly by running `pip install -r requirements.txt`.
- **FileNotFoundError:** Make sure the input Word document is placed correctly in the specified directory.
- **OSError:** The script will create necessary output directories if they do not exist, but double-check your file paths if you encounter issues.

## Contributing

We welcome contributions to improve this project! Please follow these steps to contribute:

1. Fork the repository.
2. Create a new branch: `git checkout -b feature-branch-name`.
3. Commit your changes: `git commit -m 'Add some feature'`.
4. Push to the branch: `git push origin feature-branch-name`.
5. Open a pull request.

## License

This project is licensed under the MIT License. See the `LICENSE` file for details.

## Acknowledgements

- Inspired by projects that focus on data extraction and automation.
- Thanks to the open-source community for their valuable tools and libraries.

## Contact

For any questions or issues, feel free to open an issue on GitHub or reach out to [your email here].
