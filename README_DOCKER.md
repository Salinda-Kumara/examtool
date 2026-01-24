# Deploying ExamTool to Ubuntu Server

## Prerequisites
- **Docker** must be installed on your Ubuntu server.
- **Git** (optional, if using git to transfer files).

## Step 1: Transfer Files
Copy the following files to your Ubuntu server (e.g., into a folder named `examtool`):
- `app.py`
- `Dockerfile`
- `requirements.txt`
- `.dockerignore`
- `deploy.sh` (created in the next step)
- Your Excel files (`report.xls`, etc.)

You can use `scp` (Secure Copy) from your Windows terminal:
```powershell
# Example command (replace user and ip)
scp -r "C:\Users\salu\Desktop\examtool 2\examtool\*" user@192.168.1.10:/home/user/examtool/
```

## Step 2: Run Deployment Script
On your Ubuntu server, navigate to the folder and run the deployment script:

```bash
cd ~/examtool
chmod +x deploy.sh
./deploy.sh
```

## Step 3: Access the App
Open your browser and navigate to:
`http://<your-ubuntu-server-ip>:5000`
