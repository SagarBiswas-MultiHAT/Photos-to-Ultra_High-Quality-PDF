# Ultra-High Quality Converter 

## Photo/s to PDF/s && PDF/s to Photo/s

This is a simple Windows desktop app that turns your photo(s) into a PDF(s) and PDF(s) into a photo(s).

<div align="center">

![](https://imgur.com/HF2Iufb.png)

![](https://imgur.com/YfSs9at.png)

</div>

Think of it like this:

- You pick one or more photos.
- You choose where to save.
- The app makes a PDF (one page per photo).

It tries to keep your photos looking as good as possible.

## What you need (easy)

- A Windows PC
- Python installed (if you can run the app, you already have it)
- This project folder (the one that contains:
  - `Converter-Photos to PDF_PDF to photos.py`
  - `requirements.txt`
  - zipped `venv.rar` file)

## How to run it

### Install the libraries and run

If you do NOT have `.venv` (or it doesn’t work), do this:

1. Open PowerShell in the project folder

```powershell
python -m venv venv
```

```powershell
.\venv\Scripts\Activate.ps1
```

2. Install the needed libraries:

```powershell
pip install -r requirements.txt
```

3. Start the app:

```powershell
python -u "Converter-Photos to PDF_PDF to photos.py"
```

## How to use the app

1. Click **Add…**

   - Select one or many photos.

2. (Optional) Arrange the pages

   - **Up / Down** changes the page order.
   - **Remove** deletes selected photos from the list.
   - **Clear** removes all.
   - **Sort A→Z** sorts by filename.

3. Choose how you want to save

   - **Single PDF (multi-page)** = one PDF with many pages.
   - **One PDF per photo** = many PDF files.

4. Pick where to save

   - Use **Browse…** to choose the output file/folder.

5. (Optional) Choose “Quality & page sizing”

   - If you don’t know what these mean, you can keep the defaults.

6. Click **Convert**
   - You will see progress with auto_dip value.
   - Click **Cancel** if you want to stop.

## PDF → Photos (PDF pages to images)

The app can also convert a PDF into photos (one image per page).

1. Click **PDF → Photos…**
2. Click **Add…** and choose one or more PDF files
3. Choose an **Output folder**
4. Pick **Render DPI** (300 recommended; 600 for extra sharp)
5. Pick **Image format**:

- `png` = lossless (bigger files)
- `jpg` = smaller files (uses JPEG quality setting)

6. Click **Convert**

Output filenames look like:

- `yourpdf_page_001.png`, `yourpdf_page_002.png`, ...

## Features (what every button/option really does)

### Export

- **Single PDF (multi-page)**

  - Makes one PDF file.
  - Each photo becomes a page.

- **One PDF per photo**
  - Makes many PDF files.
  - Each photo becomes its own PDF.

### Auto-rotate (EXIF)

Phones often save a photo sideways and also save a “rotate me” note.
This option reads that note and fixes the rotation.

### DPI (important, but simple)

DPI does NOT magically add detail.
It mostly decides how big the photo will be on the PDF page.

Example:

- Same photo + higher DPI = the photo looks physically smaller on the page.
- Same photo + lower DPI = the photo looks physically bigger on the page.

Here is a simple guide:

<div align="center">

| Use case                   | Best DPI        |
| -------------------------- | --------------- |
| Books / PDFs / assignments | **300**         |
| Notes, handwritten scans   | **600**         |
| Photos (normal print)      | **300**         |
| Diagrams, line art         | **600**         |
| Web / screen only          | **300 or less** |

</div>

Simple rule:

- If you’re unsure → pick **300 DPI**.
- Use **600 DPI** only when you know you need extra detail.

### Page size modes

- **auto_dpi** (good default)

  - The app tries to read DPI from your photo.
  - If your photo doesn’t have DPI information, it uses the “DPI (fallback/manual)” box.

- **dpi**

  - You choose the DPI yourself.
  - Common choices: 300 (good), 600 (very sharp for printing).

- **match_pixels**

  - Uses the photo pixel size directly as the PDF page size.
  - This can make huge pages for large photos.

- **a4 / letter**
  - Forces the PDF page to be A4 or Letter.
  - Your photo is fit inside the page.

### Embedding (this is the real “quality” choice)

This controls how the photo is put into the PDF.

- **jpeg_high** (default)

  - The app re-saves the photo as a high-quality JPEG.
  - Good when you want smaller file sizes, but still very high quality.
  - You can change **JPEG quality** (like 95, 98, 100).

- **keep_original**

  - Tries to keep the original photo file data without recompressing.
  - Best when your photo is already a nice JPEG/PNG and you want “don’t touch it”.

- **lossless_png**
  - Converts everything to PNG (no JPEG compression).
  - Biggest PDF files, but very faithful.

### Margin (pt)

- Margin is empty space around the photo.
- `0` means edge-to-edge.

### PDF metadata

If you turn on **Set PDF metadata**, you can add:

- Title
- Author

This is like the “name tag” inside the PDF.

## “Best quality” guarantee (what the app really does)

The app prefers a library called `img2pdf`.

- If `img2pdf` is installed:
  - The app can embed images into PDF in a very high-quality way (often lossless for JPEG/PNG).
- If `img2pdf` is NOT installed:
  - The app still works, but it uses a fallback method.
  - The app will show a warning so you know.

## Supported photo types

Common ones are supported:

- PNG, JPG/JPEG, BMP, GIF, TIF/TIFF, WEBP

## Troubleshooting (quick fixes)

### 1) “PowerShell won’t activate .venv”

Sometimes Windows blocks scripts.
Try running PowerShell as Administrator and then:

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

Then activate again:

```powershell
.\.venv\Scripts\Activate.ps1
```

### 2) “It says img2pdf is missing”

Run:

```powershell
pip install -r requirements.txt
```

### 2b) “PDF → Photos says PDF engine not installed”

Install dependencies:

```powershell
pip install -r requirements.txt
```

(This installs `PyMuPDF`, which the PDF → Photos feature needs.)

### 3) “The PDF looks blurry”

Most of the time the photo itself is low resolution.
Try:

- Use a better original photo, OR
- Use **lossless_png** (bigger files), OR
- Use **jpeg_high** and set **JPEG quality** to 98–100.

### 4) “The page size is weird”

Try:

- Set **Page size** to `a4` or `letter`, OR
- Use `dpi` and choose 300.

## Privacy

This app runs on your computer.
It does not upload your photos anywhere.
