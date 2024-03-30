# program-spec-generator

program-spec-generator is a docs RPA , helpful for generating Program Specifications.

## Main File Usage

### Input Folder

| File Name                           | Description                                                                             |
|-------------------------------------|-----------------------------------------------------------------------------------------|
| **[Program Spec Master Data.xlsx]** | Master Input Data                                                                       |
| **[Swagger\\\*-swagger-\*.yaml]**   | Each Project Swagger file                                                               |
| **[Logical Sequences\\\*.docx]**    | Each API Logical Sequences Section                                                      |
| **[DB-Schema.xlsx]**                | DB Schema                                                                               |
| **[Access Control.xlsx]**           | Each API Business Rules & Logic (Description / Pseudo Code) - Access Control Table Data |
| **[Validation Rules.xlsx]**         | Each API Validation Rules Table Data                                                    |

### Template Folder

| File Name                                 | Description                                                     |
|-------------------------------------------|-----------------------------------------------------------------|
| **[master_data_template.xlsx]**           | Master data format , mapping the prog_spec_template.docx        |
| **[prog_spec_excel_template.xlsx]**       | Mapping the **[prog_spec_word_template_part\*.docx]**           |
| **[prog_spec_word_template_part\*.docx]** | Base on refer/Program Specification Service_API_v0.1_DRAFT.docx |
| **[prog_spec_output_template.docx]**      | Final output.docx template content                              |
| **[prog_spec_output_raw_template.docx]**  | Final output_raw.docx template content                          |

### Transit Folder

- Temporary files / Can be used as a distribution file and collaborate with teammate.

### Python Script

#### creator.py

Mandatory Process

- Create distribute Module Folder and API File(.xlsx)

#### requester.py

Optional Process

- Semi-automatic get each API Sample Output
- Fill into Master Data File(Sample Output Sheet)
- Manually review data and fill into Master Data File(MASTER Sheet)

#### extractor.py

Mandatory Process

- Extract data from each data source
- Fill into target distribute file(.xlsx)

#### converter.py

Mandatory Process

- Compose distribute file(.xlsx) and convert to output file(.docx)

#### main.py

Compose the mandatory process

## Requirements

- Python 3.9+
- PIP

- `pip install -r requirements.txt`

```shell
docxcompose==1.4.0
numpy==1.25.2
openpyxl==3.1.2
pandas==2.0.3
python_docx==0.8.11
PyYAML==5.4
```

## Sample Run

- Prepare the required input files
- Then run follow cmd

    - run stage cmd
      ```shell
      python creator.py
      python extractor.py
      python converter.py
      ```

    - or run compose cmd

      ```shell
      python main.py
      ```
- Finally, get the output file and global search `<checkme>` marker to manually handle it.

## Workflow

![](./assets/workflow.svg)