# QPapers — College Question Paper Portal

A Flask web app for sharing question papers with Excel-based student authentication
and an admin approval workflow.

## Quick Start

### 1. Install dependencies
```bash
pip install -r requirements.txt
```

### 2. Set up environment variables
Edit `.env` (already configured with your Cloudinary keys):
```
SECRET_KEY=your-secret-key
CLOUDINARY_CLOUD_NAME=dxe1zhc8i
CLOUDINARY_API_KEY=614693399321541
CLOUDINARY_API_SECRET=I99xuKKm8tdcwUTlWg2V3qVlZ0U
ADMIN_USERNAME=admin
ADMIN_PASSWORD=admin123
```

### 3. Place Excel file
Put `ECE.xlsx` in the project root (same folder as `app.py`).

### 4. Run
```bash
python app.py
```
Open http://127.0.0.1:5000

---

## Login Details

### Student Login
- **Username:** Register Number (e.g. `2503617810621002`)
- **Password:** Same as Register Number
- Only students with `Status = REGULAR` in the Excel file can log in.

### Admin Login
- URL: `/admin/login`
- **Username:** `admin`
- **Password:** `admin123`

---

## Feature Summary

| Feature | Details |
|---|---|
| Student auth | Excel-based (pandas), no registration form needed |
| Upload | PDF only, ≤10 MB, sent to Cloudinary |
| Status flow | pending → approved / rejected |
| Student view | Only **approved** papers visible |
| Admin panel | Approve / Reject / Delete per paper |
| Filter | By department, year, subject |
| Download | Cloudinary `fl_attachment` forced download |

---

## Routes

| Route | Access |
|---|---|
| `/login` | Student login |
| `/dashboard` | Student dashboard (approved papers) |
| `/papers` | Browse approved papers |
| `/upload` | Upload PDF (logged-in students) |
| `/admin/login` | Admin login |
| `/admin/dashboard` | Admin overview |
| `/admin/papers?status=pending` | Filter papers by status |
| `/approve/<id>` | POST — approve paper |
| `/reject/<id>` | POST — reject paper |
| `/delete/<id>` | POST — delete paper |
