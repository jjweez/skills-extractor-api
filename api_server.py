from fastapi import FastAPI, UploadFile, Form
from fastapi.responses import JSONResponse
from skills_extractor import process_skills
import shutil
import os
import uuid
import traceback

app = FastAPI()

@app.post("/extract")
async def extract(
    file: UploadFile,
    sheet: str = Form(...),
    client: str = Form(...),
    sender: str = Form(...)
):
    try:
        # Save uploaded file temporarily (local temp dir)
#        temp_filename = f"C:/Users/jared/AppData/Local/Temp/tmp_{uuid.uuid4()}_{file.filename}"
        import tempfile

        # Use a universal temp directory that works on Windows, Mac, and Linux
        temp_dir = tempfile.gettempdir()
        temp_filename = os.path.join(temp_dir, f"tmp_{uuid.uuid4()}_{file.filename}")
 
        with open(temp_filename, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        # Run extractor using the temp file
        result = process_skills(temp_filename, sheet, client, sender)

        # Move review file back beside original, but clean the name
        review_path = result["output_file"]
#        original_dir = "C:/Users/jared/OneDrive/_PTP"
        # Save output next to input if local, otherwise current working dir
        if os.name == "nt":  # Windows
            original_dir = "C:/Users/jared/OneDrive/_PTP"
        else:
            original_dir = os.getcwd()


        if os.path.exists(review_path):
            base_name = os.path.basename(file.filename)
            review_name = os.path.splitext(base_name)[0] + "_review.xlsx"
            dest_path = os.path.join(original_dir, review_name)
            shutil.move(review_path, dest_path)
            result["output_file"] = dest_path

        # Close file handle
        try:
            file.file.close()
        except Exception:
            pass

        # Make message more readable
        raw_message = result["message"]
        formatted_message = (
            raw_message.replace("\\n", "\n")
                       .replace("\r\n", "\n")
                       .strip()
        )

        # Build HTML version for GPT or email display
        html_message = (
            "<p>" + formatted_message.replace("\n\n", "</p><p>")
                                     .replace("\n", "<br>") + "</p>"
        )

        result["message"] = formatted_message
        result["message_html"] = html_message

        # Console preview
        print("\n" + "=" * 60)
        print("📬  SHARE MESSAGE PREVIEW:")
        print("-" * 60)
        print(formatted_message)
        print("=" * 60 + "\n")

        # Return JSON
        return JSONResponse(content={"result": result})

    except Exception as e:
        print("❌ ERROR IN /extract ENDPOINT ❌")
        traceback.print_exc()
        return JSONResponse(content={"error": str(e)}, status_code=500)
