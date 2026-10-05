# Product CSV Comparator — Render Deployment

## Deploy with Render

1. Put this project in a GitHub repository.
2. Go to https://render.com/ and sign in.
3. Choose **New → Web Service**.
4. Connect the GitHub repository.
5. Use:
   - Runtime: Python
   - Build Command: `pip install -r requirements.txt`
   - Start Command: `gunicorn app:app`
6. Create the service.
7. Render will provide a public `onrender.com` URL.

The included `render.yaml` can also be used for the service configuration.

## Local run

```bash
pip install -r requirements.txt
python app.py
```

Open http://127.0.0.1:5000
