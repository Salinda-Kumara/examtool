# Deploying ExamTool to Ubuntu Server

## Prerequisites
- **Docker** installed on your Ubuntu server.
- **Git** installed on your Ubuntu server.

## Step 1: Clone the Repository
On your Ubuntu server, clone the repository from GitHub:

```bash
# Clone the repository (replace URL with your actual repo URL if different)
git clone https://github.com/Salinda-Kumara/examtool.git
cd examtool
```

## Step 2: Transfer Data Files
Since Excel files and PDFs are ignored by Git (via `.gitignore`), you need to copy them manually to the `examtool` directory on your server:
- `report.xls`
- `report.pdf`

You can use `scp` from your Windows machine:
```powershell
scp "C:\Users\salu\Desktop\examtool 2\examtool\report.xls" user@<server-ip>:~/examtool/
```

## Step 3: Run Deployment Script
Make the script executable and run it:

```bash
chmod +x deploy.sh
./deploy.sh
```

## Step 4: Access the App
Open your browser and navigate to:
`http://<your-ubuntu-server-ip>:5000`
