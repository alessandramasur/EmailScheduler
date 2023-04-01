# EmailScheduler
Python script to generate and schedule E-Mails via the Outlook App. After excecuting the script the scheduled E-mails will appear in the outbox folder and are sent automatically at the scheduled time. The Python script takes a CSV file with E-Mail addresses and a YAML file with the email subject and body as input and drafts the same E-Mail to all recipients. A variable number of E-Mails are scheduled to be sent every 10 minutes, from 7:00am to 6:50pm, starting from the next day. Duplicate E-Mail addresses from the input are deleted. All recipients and their scheduled delivery time are saved in `scheduled_mails.CSV`.

E-Mails are scheduled every 10 minutes to ensure the sender's E-Mail account will not be blocked. This script is intended to be used to send a very large number of E-Mails.

## Command line arguments

- `--inputCSV`: Path to the input CSV file. **Required.**
- `--delimiter`: Delimiter string/symbol in CSV file. 
    Default = ";"
- `--inputYAML`: Path to the input YAML file. 
    Default = "mailcontent.yaml"
- `--mailsPerTenMin`: How many mails should be sent every ten minutes. 
    Default = "35"
- `--previewMode`: Set 'True' or 'False' for turning preview mode on or off. If preview mode is turned on, only first E-Mail is generated for preview. No E-Mails are sent or scheduled in preview mode. 
    Default = "True"

## Input files
For better overview input files are in `input` folder.

- ##### input CSV file: 
    For email addresses and surnames. First column in CSV file contains E-Mail addresses, second column contains respective surname, third column contains `science` or `admin` to differenciate which E-Mail body should be sent.
    Example row in CSV file:  `exampleemail@eg.com;Smith`
- ##### input YAML file:
    For the email subject and body. E-Mail body should be in HTML format. Should contain surname placeholders `{name}` in which surnames from CSV files will be inserted. Default file name is "mailbody.yaml". Keys should be called `subject` and `body`.
    Example template:
    ```YAML
    ---
    # input for the mail subject and mail body
    subject: "Example subject"
    body: "Hello {name},<br><br>this is an example email."

    ```
## Output file
All recipients and their scheduled delivery time are saved in a file `scheduled_mails.CSV` in the same directory as the Python script.

## Command line execution:
Example execution in command line:

```bash
python emailscheduler.py --inputCSV input/recipients.CSV
```
```bash
python emailscheduler.py --inputCSV input/recipients.CSV --previewMode False
```

## Install requirements

```bash
pip install -r requirements.txt
```