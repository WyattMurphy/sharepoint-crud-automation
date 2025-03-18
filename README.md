# **SharePoint CRUD Automation**

## **Overview**
This repository provides an automated framework for performing CRUD (Create, Read, Update, Delete) operations on SharePoint files. It enables seamless file management directly from Python, with a focus on minimizing redundancy and maximizing reusability. The system allows for scheduled or manual automation of file uploads, downloads, deletions, and listings to streamline corporate data processes.

### **Objectives:**
- Automate SharePoint file operations to replace manual processes.
- Centralize connection logic to avoid redundant authentication and session management.
- Provide modular, reusable, and easily extendable scripts for common SharePoint operations.
- Prepare the infrastructure for further automation, such as scheduling tasks and integrating into larger workflows.

## **File Structure**

```
sharepoint-crud-automation/
├── automation/                           # Automation-related scripts
│   ├── upload_file.py                    # Script to automate file uploads to SharePoint
│   ├── download_file.py                  # Script to automate file downloads from SharePoint
│   ├── delete_file.py                    # Script to automate file deletions in SharePoint
│   ├── list_files.py                     # Script to automate file listings in SharePoint
│   └── scheduler/                        # Scheduler for periodic tasks
│       └── job_scheduler.py              # Automation of scheduled tasks
├── src/                                  # Core functionalities
│   ├── sharepoint_client.py              # SharePoint client and connection handling
│   ├── file_operations.py                # Functions for CRUD operations
│   └── logger.py                         # Logging configuration and utilities
├── config/                               # Configuration files
│   ├── config.yaml                       # General configuration (SharePoint URL, credentials)
│   └── credentials.json                  # Secure credentials for SharePoint (do not commit)
├── data/                                 # Data storage (temporary for automation tasks)
│   ├── raw/                              # Raw data (before processing)
│   ├── processed/                        # Processed data (after automation or cleaning)
│   └── logs/                             # Logs generated during automation tasks
├── notebooks/                            # Jupyter Notebooks for exploratory tasks
├── requirements.txt                      # Project dependencies
├── README.md                             # Project documentation
└── .gitignore                            # Files to ignore in version control
```

## **Key Components**

1. **SharePointClient Class** (`src/sharepoint_client.py`):
   - Centralizes the connection logic for SharePoint.
   - Manages authentication and session handling to ensure efficient resource use.

2. **CRUD Operations**:
   - **Upload**: Automates file uploads to SharePoint (`automation/upload_file.py`).
   - **Download**: Automates file downloads from SharePoint (`automation/download_file.py`).
   - **Delete**: Automates file deletions on SharePoint (`automation/delete_file.py`).
   - **List**: Retrieves and lists files in a specified SharePoint folder (`automation/list_files.py`).

3. **Scheduler** (`automation/scheduler/job_scheduler.py`):
   - Manage scheduled tasks to automate operations like periodic file uploads or downloads.

4. **Configuration**:
   - Configuration files store essential parameters such as SharePoint URLs, credentials, and automation settings (`config/config.yaml`).

5. **Logs**:
   - Logs are generated to monitor the success of operations and catch errors (`data/logs/`).

## **Usage**

### 1. **Clone the Repository**
```bash
git clone https://github.com/yourusername/sharepoint-crud-automation.git
cd sharepoint-crud-automation
```

### 2. **Install Dependencies**
```bash
pip install -r requirements.txt
```

### 3. **Set Up Configuration**
- Modify `config/config.yaml` to include your SharePoint site URL and other configurations.
- Ensure your SharePoint credentials are stored securely in `config/credentials.json` (avoid committing sensitive data).

### 4. **Run Operations**

You can run each operation individually. For example:

```bash
python automation/upload_file.py --file_path "local_file.xlsx" --target_path "Documents/Shared"
```

### 5. **Automated Workflow**
You can create a custom workflow that combines multiple operations. Example:

```python
from automation.upload_file import upload_file
from automation.list_files import list_files
from automation.download_file import download_file

def automated_workflow():
    upload_response = upload_file("local_file.xlsx", "Documents/Shared", site_url, username, password)
    file_list = list_files("Documents/Shared", site_url, username, password)
    download_response = download_file("uploaded_file.xlsx", "Documents/Shared", site_url, username, password)
    return upload_response, file_list, download_response
```

### 6. **Scheduling Tasks**
You can schedule operations with `automation/scheduler/job_scheduler.py` 

## **Future Work**
- Add advanced features like error handling, retry mechanisms, and automated reporting.
- Integrate with other data processing systems, like Excel automation or database updates.

---

This setup ensures a **modular, reusable, and efficient** SharePoint automation pipeline while laying the groundwork for future expansion and integration into larger business processes.

