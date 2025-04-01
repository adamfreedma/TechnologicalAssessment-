# Socio_and_Sagabz

## Features
- **Data Preprocessing**: Cleans and structures input Excel files for further processing.
- **Classification**: Uses AI-based classification for evaluation comments.
- **Statistical Analysis**: Computes means, standard deviations, and sigma values.
- **Report Generation**: Creates personalized Word reports with graphical representations of performance metrics.
- **Historical Comparison**: Allows integration of past evaluation data for trend analysis.

## Project Structure
```
Socio_and_Sagabz/
│── main.py                    # Entry point for running socio/sagabz analysis
│── pre_process.py             # Data preprocessing module
│── run_obj.py                 # Main class to execute the evaluation pipeline
│── run_sagabz_socio.py        # Handles socio and sagabz execution
│── sagabz_socio_pre_process.py # Preprocessing for both socio and sagabz
│── socio_to_classification.py  # AI-based classification module
│── statistics_socio.py        # Computes statistical metrics
│── column_constants.py        # Column name mappings and constants
│── docx_helper.py             # Word document generation and report generation
│── docx_sagabz_socio.py       # Report utilities for socio and sagabz
```

## Dependencies
To run the system, the following Python packages are required:
- `pandas`
- `numpy`
- `matplotlib`
- `networkx`
- `openpyxl`
- `python-docx`
- `tqdm`
- `dotenv`
- `google-generativeai`
- `docxtpl`

To install all dependencies, run:
```sh
pip install -r requirements.txt
```

## Running the System
1. **Setup API key for the classification (only for classification)** get Google API key for using the gemeni model. add it as an environment variable `GOOGLE_API_KEY=`
2. **Prepare Data**: Place input Excel files in the `Excels/` directory.
2. **Run the Main Analysis**:
   - To execute socio evaluation:
     ```sh
     python main.py --run-socio --raw-data-path "Socio_and_sagabz/Excels"
     ```
   - To execute sagabz evaluation:
     ```sh
     python main.py --run-socio --run-sagabz
     ```
3. **Optional Parameters**:
   - To skip classification:
     ```sh
     python main.py --run-socio --dont-run-classification
     ```
   - To do one task like *split_excel*, *statistics* or *word_build*:
     ```sh
     python main.py --run-socio --start-task word_build
     ```
   - To start the run from a specific cadet
     ```sh
     python main.py --run-socio --start-cadet "name of cadet"
     ```
   - To include the data of last semester:
     ```sh
     python main.py --run-socio --old-stats-path "path to old stats.xlsx file"
     ```
   - To make censor the names of the cadets in the outputs:
     ```sh
     python main.py --run-socio --names-to-hashes
     ```

## Output
- **Excel Reports**: Contains structured evaluation metrics.
- **Word Reports**: Personalized reports for individuals.

## New Fromat:
- in the doc itself, you can insert tags, s.a ```{{tag}}```.
- to replace it with a socio element you can call ```find_paragraph_by_tag()``` or ```find_table_cell_by_tag()``` clear it and add you element.
- ```create_word_file()``` creates a word file for a single person, it calls subfunctions that use the functions mentioned above to insert the elements into the file.
- ```run_word_creation()``` iterates over all names and creates a word file for each one.