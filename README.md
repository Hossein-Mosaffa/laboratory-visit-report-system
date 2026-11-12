# Laboratory Visit Report System

A web-based application built with Python, Flask, and SQLite for recording and managing laboratory visit and evaluation reports.

The system allows an administrator to manage laboratory information, record evaluation details, filter records, and export filtered results as an Excel file.

## Features

- Add, view, and edit laboratory information
- Record laboratory visit and evaluation details
- Store evaluation results, non-conformance information, follow-up dates, and committee decisions
- Filter records by laboratory name, industry, certificate issue date, certificate validity date, and evaluation date
- Export filtered records to an Excel `.xlsx` file
- Simple admin login system
- SQLite database with no external database server required
- Persian-language user interface

## Technologies

- Python 3
- Flask
- SQLite
- Pandas
- OpenPyXL
- HTML, CSS, and JavaScript

## Project Structure

```text
laboratory-visit-report-system/
│
├── app.py
├── requirements.txt
├── README.md
├── .gitignore
│
├── templates/
│   ├── login.html
│   ├── admin.html
│   ├── records.html
│   ├── record_form.html
│   ├── edit_record.html
│   ├── laboratory_form.html
│   ├── edit_laboratory.html
│   ├── show_labs.html
│   └── report.html
│
└── static/
    └── css/
        └── style.css
```

## Database

The application automatically creates an SQLite database file named `records.db` when it starts.

It has two main tables:

| Table | Purpose |
|---|---|
| `laboratories` | Stores laboratory name, industry, certificate number, and first certificate issue date |
| `records` | Stores visit and evaluation information linked to a laboratory |

Each record includes:

- Certificate validity date
- Evaluation type and date
- Evaluation team members
- Evaluation result
- Number of non-conformities
- Deadline for correcting non-conformities
- Follow-up evaluation date
- File submission and technical committee dates
- Committee result
- Technical manager
- Registrant

## Installation

1. Clone the repository:

```bash
git clone [https://github.com/your-username/laboratory-visit-report-system.git](https://github.com/your-username/laboratory-visit-report-system.git)
cd laboratory-visit-report-system
```

2. Create and activate a virtual environment:

```bash
python -m venv venv

# Windows
venv\Scripts\activate

# Linux / macOS
source venv/bin/activate
```

3. Install the required packages:

```bash
pip install -r requirements.txt
```

4. Run the application:

```bash
python app.py
```

5. Open the application in your browser:

```text
http://127.0.0.1:5001
```

## Admin Access

The application includes a basic admin login system.

For security, admin credentials should be stored in environment variables rather than hardcoded in the source code.

## Excel Report Export

After applying filters on the report page, the results can be downloaded as an Excel file.

The exported report includes laboratory details, evaluation information, non-conformance data, committee results, and responsible personnel.

## Future Improvements

- Add user roles and permissions
- Use hashed passwords for authentication
- Add record deletion with confirmation
- Add search and pagination
- Add PDF report generation
- Add REST API endpoints
- Add automated tests
- Add Docker support

## License

This project is licensed under the MIT License.

[Hossein Mosaffa's personal website](https://hossein-mosaffa.github.io)
