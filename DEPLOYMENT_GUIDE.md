# Deploying Your Streamlit App Online  
  
## Option 1: Streamlit Sharing (Recommended for beginners)  
  
### Step-by-step instructions:  
  
1. **Create a GitHub repository** for your project:  
   - Go to https://github.com and create a new repository  
   - Name it something like "e-class-dbf-updater"  
   - Don't initialize with README or .gitignore yet  
  
2. **Prepare your files**:  
   - Make sure your `app.py` and `requirements.txt` are in the repository root  
   - Your project should have this structure:  
     ```  
     e-class-dbf-updater/  
     |^-- app.py  
     |^-- requirements.txt  
     ^^^^^^^^^^^^^-- .streamlit/  
         ^^^^^^^^^^^^^^^^^^^^^^^^^^^-- config.toml  
     ```  
  
3. **Push your code to GitHub**:  
   ```bash  
   git init  
   git add .  
   git commit -m "Initial commit"  
   git branch -M main  
   git remote add origin YOUR_GITHUB_REPOSITORY_URL  
   git push -u origin main  
   ```  
  
4. **Deploy to Streamlit Sharing**:  
   - Go to https://share.streamlit.io  
   - Sign in with your GitHub account  
   - Click "New app"  
   - Select your repository  
   - Choose the main branch  
   - Set the file path to `app.py`  
   - Click "Deploy!"  
  
## Option 2: Heroku Deployment  
  
1. **Create a Procfile**:  
   ```  
   web: streamlit run app.py --server.port=$PORT  
   ```  
  
2. **Add runtime.txt** (optional):  
   ```  
   python-3.9  
   ```  
  
3. **Deploy using Heroku CLI**:  
   ```bash  
   heroku create your-app-name  
   git push heroku main  
   ```  
  
## Option 3: Railway Deployment  
  
1. Add a `runtime.txt` file:  
   ```  
   python-3.9  
   ```  
  
2. Connect your GitHub repo to Railway.app  
3. Railway will automatically detect and deploy your Python/Streamlit app  
  
## Option 4: Render Deployment  
  
1. Add a `render.yaml` file:  
  ```yaml  
  services:  
    - type: web  
      name: e-class-dbf-updater  
      env: python  
      region: oregon  
      buildCommand: pip install -r requirements.txt  
      startCommand: streamlit run app.py --server.port=$PORT --server.address=0.0.0.0  
      envVars:  
        - key: PYTHON_VERSION  
          value: 3.9.16  
  ```  
  
## Important Notes:  
  
- Your `requirements.txt` must be at the root of the repository  
- The main Streamlit file must be named `app.py` (as you already have)  
- Make sure all dependencies in `requirements.txt` are compatible with the deployment platform  
- The app will be publicly accessible once deployed  
  
## Testing locally before deployment:  
  
Run your app locally to make sure everything works:  
```bash  
pip install -r requirements.txt  
streamlit run app.py  
```  
  
## Troubleshooting Common Issues:  
  
1. **Module not found errors**: Ensure all dependencies are in requirements.txt  
2. **Port binding issues**: On most platforms, use the PORT environment variable  
3. **File upload limits**: Different platforms have different limits for file uploads  
4. **Memory limits**: Free tiers often have memory limitations for processing large files  
  
