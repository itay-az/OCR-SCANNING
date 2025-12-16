"""
מודול לעיבוד קבצי PDF - קריאת טקסט, OCR, וחיפוש REGEX
"""
import os
import re
import shutil
import fitz  # PyMuPDF
import importlib.util, pkgutil
if not hasattr(pkgutil, "find_loader"):
    pkgutil.find_loader = lambda name: importlib.util.find_spec(name)
 
import pytesseract
from PIL import Image
import io
import pytesseract

# === הגדרות TESSERACT ===
DEFAULT_TESSERACT_PATH = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
if os.path.exists(DEFAULT_TESSERACT_PATH):
    pytesseract.pytesseract.tesseract_cmd = DEFAULT_TESSERACT_PATH

def is_valid_israeli_id(id_str):
    """
    בודק תקינות תעודת זהות ישראלית (אלגוריתם ה-Luhn/ספרת ביקורת)
    """
    # ניקוי רווחים ומקפים אם השתרבבו
    id_str = str(id_str).strip().replace('-', '')
    
    if not id_str.isdigit():
        return False
    if len(id_str) > 9:
        return False
    if len(id_str) < 9:
        # ריפוד באפסים אם המספר קצר מ-9 (למרות שה-Regex שלנו מחפש 9)
        id_str = id_str.zfill(9)

    total = 0
    for i in range(9):
        val = int(id_str[i])
        # הכפלה ב-1 או ב-2 לסירוגין
        if i % 2 == 1:
            val *= 2
        
        # אם התוצאה דו-ספרתית, מסכמים את הספרות (למשל 14 -> 1+4=5)
        if val > 9:
            val = (val // 10) + (val % 10)
            
        total += val

    # בדיקה אם הסכום מתחלק ב-10 ללא שארית
    return (total % 10) == 0

def extract_text_from_pdf(pdf_path):
    """מנסה לחלץ טקסט ישירות מקובץ PDF searchable"""
    try:
        doc = fitz.open(pdf_path)
        text = ""
        for page_num in range(len(doc)):
            page = doc[page_num]
            text += page.get_text()
        doc.close()
        return text.strip()
    except Exception as e:
        print(f"שגיאה בקריאת PDF {pdf_path}: {e}")
        return ""

def pdf_to_images(pdf_path):
    """ממיר קובץ PDF לרשימת תמונות באיכות גבוהה ל-OCR"""
    try:
        doc = fitz.open(pdf_path)
        images = []
        for page_num in range(len(doc)):
            page = doc[page_num]
            zoom = 300 / 72  
            mat = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=mat)
            img_data = pix.tobytes("png")
            img = Image.open(io.BytesIO(img_data))
            images.append(img)
        doc.close()
        return images
    except Exception as e:
        print(f"שגיאה בהמרת PDF לתמונות {pdf_path}: {e}")
        return []

def save_images_to_pdf(images, output_path):
    """שומר רשימת תמונות כקובץ PDF (לשמירה מחדש אחרי סיבוב)"""
    if not images:
        return
    try:
        converted_images = []
        for img in images:
            if img.mode != 'RGB':
                converted_images.append(img.convert('RGB'))
            else:
                converted_images.append(img)
        
        converted_images[0].save(
            output_path, "PDF", resolution=300, save_all=True, append_images=converted_images[1:]
        )
    except Exception as e:
        print(f"שגיאה בשמירת PDF מתוקן: {e}")

def perform_ocr_on_images(images):
    """מבצע OCR על רשימת תמונות"""
    full_text = ""
    custom_config = r'--oem 3 --psm 6'
    
    try:
        for img in images:
            gray_img = img.convert('L')
            try:
                text = pytesseract.image_to_string(gray_img, lang='heb+eng', config=custom_config)
            except Exception:
                text = pytesseract.image_to_string(gray_img, lang='eng', config=custom_config)
            full_text += text + "\n"
    except Exception as e:
        print(f"שגיאה ב-OCR: {e}")
    
    return full_text.strip()

def find_regex_match(text, regex_pattern):
    """
    מחפש התאמה בטקסט.
    שינוי קריטי: עובר על כל ההתאמות ובודק תקינות תעודת זהות.
    """
    try:
        # משתמשים ב-finditer כדי לעבור על כל התוצאות האפשריות
        matches = re.finditer(regex_pattern, text, re.MULTILINE)
        
        for match in matches:
            candidate = match.group(0).strip()
            
            # בדיקה: האם זה מספר תקין מבחינת ספרת ביקורת?
            # הערה: אם אתה מחפש משהו שאינו ת"ז (למשל סתם קטלוג), אפשר לבטל את הבדיקה הזו
            if is_valid_israeli_id(candidate):
                return candidate
            else:
                print(f"דלג על מספר לא תקין (ספרת ביקורת שגויה): {candidate}")
                
        return None
    except re.error as e:
        print(f"שגיאה בתבנית REGEX: {e}")
        return None

def process_pdf_file(pdf_path, regex_pattern):
    """הפונקציה הראשית לעיבוד קובץ"""
    text = extract_text_from_pdf(pdf_path)
    images = []
    
    if not text or len(text.strip()) < 5: 
        try:
            images = pdf_to_images(pdf_path)
            if images:
                ocr_text = perform_ocr_on_images(images)
                text += "\n" + ocr_text
        except Exception as e:
            print(f"OCR נכשל: {e}")
    
    match = None
    if text:
        match = find_regex_match(text, regex_pattern)
    
    if match:
        return match

    # ניסיון סיבוב (Rotation)
    if images:
        print(f"לא נמצאה התאמה. מנסה לסובב ב-180 מעלות...")
        try:
            rotated_images = [img.rotate(180) for img in images]
            rotated_text = perform_ocr_on_images(rotated_images)
            rotated_match = find_regex_match(rotated_text, regex_pattern)
            
            if rotated_match:
                print(f"✓ נמצאה התאמה לאחר סיבוב! שומר מחדש...")
                save_images_to_pdf(rotated_images, pdf_path)
                return rotated_match
        except Exception as e:
            print(f"שגיאה בניסיון סיבוב: {e}")

    return None

def generate_id_folder_path(root_folder, id_number):
    """יוצרת נתיב בתיקייה ייעודית"""
    id_folder_path = os.path.join(root_folder, id_number)
    
    if not os.path.exists(id_folder_path):
        try:
            os.makedirs(id_folder_path, exist_ok=True)
        except OSError as e:
            print(f"שגיאה ביצירת תיקייה {id_folder_path}: {e}")
            return get_safe_filename(root_folder, id_number, ".pdf")

    counter = 1
    while True:
        filename = f"{id_number}-{counter}.pdf"
        full_path = os.path.join(id_folder_path, filename)
        if not os.path.exists(full_path):
            return full_path
        counter += 1

def get_safe_filename(directory, base_name, extension=".pdf"):
    """שם בטוח (fallback)"""
    safe_name = re.sub(r'[<>:"/\\|?*]', '', base_name).strip()
    full_path = os.path.join(directory, safe_name + extension)
    counter = 1
    while os.path.exists(full_path):
        new_name = f"{safe_name}_{counter}"
        full_path = os.path.join(directory, new_name + extension)
        counter += 1
    return full_path

def get_next_file_number(destination_folder, id_number):
    """
    מחזירה את המספר העוקב הבא בתיקיית ID.
    סופרת את כל הקבצים הקיימים בתיקייה ומחזירה את המספר הבא.
    """
    id_folder_path = os.path.join(destination_folder, id_number)
    
    if not os.path.exists(id_folder_path):
        return 1
    
    # סופרת את כל הקבצים בתיקייה שמתחילים ב-{id_number}-
    max_number = 0
    try:
        for filename in os.listdir(id_folder_path):
            if filename.startswith(f"{id_number}-") and filename.endswith(".pdf"):
                # מנסה לחלץ את המספר מהשם
                try:
                    # פורמט: {id_number}-{number}.pdf
                    number_part = filename[len(id_number) + 1:-4]  # מוריד את {id_number}- ואת .pdf
                    number = int(number_part)
                    if number > max_number:
                        max_number = number
                except ValueError:
                    continue
    except OSError:
        pass
    
    return max_number + 1

def process_folder_with_destination(source_folder, destination_folder, regex_pattern, log_callback=None):
    """
    עיבוד תיקיית מקור והעתקת קבצים לתיקיית יעד.
    - קוראת רק טקסט מ-searchable PDF (ללא OCR)
    - מחפשת תעודת זהות באמצעות REGEX
    - בודקת תקינות תעודת זהות ישראלית
    - מעתיקה קבצים לתיקיית יעד לפי תעודת זהות (או unidentified)
    - לא מוחקת/מזיזה קבצים מהמקור
    """
    if log_callback:
        log_callback(f"מתחיל עיבוד תיקיית מקור: {source_folder}")
        log_callback(f"תיקיית יעד: {destination_folder}")
    
    stats = {'success_count': 0, 'failed_count': 0, 'unidentified_count': 0, 'errors': []}
    
    # יצירת תיקיית unidentified אם לא קיימת
    unidentified_folder = os.path.join(destination_folder, "unidentified")
    try:
        os.makedirs(unidentified_folder, exist_ok=True)
    except OSError as e:
        if log_callback:
            log_callback(f"שגיאה ביצירת תיקיית unidentified: {e}")
        stats['errors'].append(f"שגיאה ביצירת תיקיית unidentified: {e}")
    
    # רשימת קבצי PDF בתיקיית המקור
    try:
        pdf_files = [f for f in os.listdir(source_folder) 
                     if f.lower().endswith('.pdf') and os.path.isfile(os.path.join(source_folder, f))]
    except OSError as e:
        if log_callback:
            log_callback(f"שגיאה בקריאת תיקיית המקור: {e}")
        stats['errors'].append(f"שגיאה בקריאת תיקיית המקור: {e}")
        return stats
    
    if log_callback:
        log_callback(f"נמצאו {len(pdf_files)} קבצי PDF לעיבוד\n")
    
    for pdf_file in pdf_files:
        pdf_path = os.path.join(source_folder, pdf_file)
        
        if log_callback:
            log_callback(f"מעבד: {pdf_file}")
        
        try:
            # קריאת טקסט מ-searchable PDF (ללא OCR)
            text = extract_text_from_pdf(pdf_path)
            
            if not text or len(text.strip()) < 5:
                if log_callback:
                    log_callback(f"   ⚠ לא נמצא טקסט בקובץ (אולי לא searchable PDF)")
                # מעתיק ל-unidentified
                dest_path = os.path.join(unidentified_folder, pdf_file)
                try:
                    shutil.copy2(pdf_path, dest_path)
                    if log_callback:
                        log_callback(f"   → הועתק ל-unidentified")
                    stats['unidentified_count'] += 1
                except Exception as e:
                    if log_callback:
                        log_callback(f"   ✗ שגיאה בהעתקה: {e}")
                    stats['errors'].append(f"שגיאה בהעתקת {pdf_file}: {e}")
                    stats['failed_count'] += 1
                continue
            
            # חיפוש תעודת זהות באמצעות REGEX
            if log_callback:
                log_callback(f"   מחפש תעודת זהות...")
            
            match_value = find_regex_match(text, regex_pattern)
            
            # בדיקה מפורשת נוספת של תקינות תעודת זהות
            if match_value:
                if log_callback:
                    log_callback(f"   נמצא מספר: {match_value}")
                
                if is_valid_israeli_id(match_value):
                    if log_callback:
                        log_callback(f"   ✓ תעודת זהות תקנית: {match_value}")
                    
                    # יצירת תיקיית ID אם לא קיימת
                    id_folder_path = os.path.join(destination_folder, match_value)
                    try:
                        os.makedirs(id_folder_path, exist_ok=True)
                    except OSError as e:
                        if log_callback:
                            log_callback(f"   ✗ שגיאה ביצירת תיקייה: {e}")
                        stats['errors'].append(f"שגיאה ביצירת תיקייה {match_value}: {e}")
                        stats['failed_count'] += 1
                        continue
                    
                    # קבלת המספר העוקב הבא
                    next_number = get_next_file_number(destination_folder, match_value)
                    new_filename = f"{match_value}-{next_number}.pdf"
                    dest_path = os.path.join(id_folder_path, new_filename)
                    
                    # העתקת הקובץ
                    try:
                        shutil.copy2(pdf_path, dest_path)
                        if log_callback:
                            log_callback(f"   ✓ הועתק ל: {match_value}/{new_filename}")
                        stats['success_count'] += 1
                    except Exception as e:
                        if log_callback:
                            log_callback(f"   ✗ שגיאה בהעתקה: {e}")
                        stats['errors'].append(f"שגיאה בהעתקת {pdf_file}: {e}")
                        stats['failed_count'] += 1
                else:
                    # נמצא מספר אבל לא תעודת זהות תקנית
                    if log_callback:
                        log_callback(f"   ⚠ נמצא מספר {match_value} אבל לא תעודת זהות תקנית")
                    # מעתיק ל-unidentified
                    dest_path = os.path.join(unidentified_folder, pdf_file)
                    try:
                        shutil.copy2(pdf_path, dest_path)
                        if log_callback:
                            log_callback(f"   → הועתק ל-unidentified")
                        stats['unidentified_count'] += 1
                    except Exception as e:
                        if log_callback:
                            log_callback(f"   ✗ שגיאה בהעתקה: {e}")
                        stats['errors'].append(f"שגיאה בהעתקת {pdf_file}: {e}")
                        stats['failed_count'] += 1
            else:
                # לא נמצאה תעודת זהות
                if log_callback:
                    log_callback(f"   ⚠ לא נמצאה תעודת זהות")
                # מעתיק ל-unidentified
                dest_path = os.path.join(unidentified_folder, pdf_file)
                try:
                    shutil.copy2(pdf_path, dest_path)
                    if log_callback:
                        log_callback(f"   → הועתק ל-unidentified")
                    stats['unidentified_count'] += 1
                except Exception as e:
                    if log_callback:
                        log_callback(f"   ✗ שגיאה בהעתקה: {e}")
                    stats['errors'].append(f"שגיאה בהעתקת {pdf_file}: {e}")
                    stats['failed_count'] += 1
                    
        except Exception as e:
            if log_callback:
                log_callback(f"   ✗ שגיאה בעיבוד: {e}")
            stats['errors'].append(f"שגיאה בעיבוד {pdf_file}: {e}")
            stats['failed_count'] += 1
    
    # סיכום
    if log_callback:
        log_callback(f"\n=== סיכום ===")
        log_callback(f"הושלמו בהצלחה: {stats['success_count']}")
        log_callback(f"לא זוהו (unidentified): {stats['unidentified_count']}")
        log_callback(f"שגיאות: {stats['failed_count']}")
        if stats['errors']:
            log_callback(f"פרטי שגיאות: {len(stats['errors'])}")
    
    return stats

def process_folder(folder_path, regex_pattern, log_callback=None):
    """עיבוד תיקייה"""
    if log_callback:
        log_callback(f"מתחיל עיבוד תיקייה: {folder_path}")
    
    stats = {'success_count': 0, 'failed_count': 0, 'errors': []}
    
    pdf_files = [f for f in os.listdir(folder_path) 
                 if f.lower().endswith('.pdf') and os.path.isfile(os.path.join(folder_path, f))]
    
    for pdf_file in pdf_files:
        pdf_path = os.path.join(folder_path, pdf_file)
        try:
            match_value = process_pdf_file(pdf_path, regex_pattern)
            
            if match_value:
                new_full_path = generate_id_folder_path(folder_path, match_value)
                if new_full_path != pdf_path:
                    os.rename(pdf_path, new_full_path)
                    new_filename = os.path.basename(new_full_path)
                    parent_folder = os.path.basename(os.path.dirname(new_full_path))
                    if log_callback: 
                        log_callback(f"✓ {pdf_file} -> {parent_folder}/{new_filename}")
                    stats['success_count'] += 1
            else:
                if log_callback: log_callback(f"✗ {pdf_file}: לא נמצאה תאמה (או מספר לא תקין)")
                stats['failed_count'] += 1
                
        except Exception as e:
            stats['errors'].append(str(e))
            stats['failed_count'] += 1
            
    return stats