"""
מודול לסריקה ישירה מסורקים ב-Windows - תמיכה ב-ADF (סריקת אצווה)
"""
import os
import win32com.client
import pythoncom
import win32api
import tempfile
from uuid import uuid4
from PIL import Image

# קבועים של WIA
WIA_MPC_HANDLE_DOCUMENT_HANDLING_SELECT = 3088
WIA_MPC_HANDLE_DOCUMENT_HANDLING_STATUS = 3087
WIA_FEEDER = 1
WIA_FLATBED = 2
WIA_DUPLEX = 4

def get_scanners():
    """מחזיר רשימה של סורקים זמינים"""
    try:
        pythoncom.CoInitialize()
        device_manager = win32com.client.Dispatch("WIA.DeviceManager")
        scanners = []
        for i in range(1, device_manager.DeviceInfos.Count + 1):
            device_info = device_manager.DeviceInfos(i)
            scanners.append(device_info.Properties("Name").Value)
        return scanners
    except Exception as e:
        print(f"שגיאה בקבלת רשימת סורקים: {e}")
        return []

def configure_adf(device):
    """מנסה להגדיר את הסורק לעבוד במצב מזין דפים (ADF)"""
    try:
        # מנסה למצוא את המאפיין שאחראי על מקור הנייר
        for prop in device.Properties:
            if prop.PropertyID == WIA_MPC_HANDLE_DOCUMENT_HANDLING_SELECT:
                # מנסה להגדיר ל-FEEDER (ערך 1)
                # אם הסורק תומך, זה יגרום לו לקחת מהמגש
                try:
                    prop.Value = WIA_FEEDER
                except:
                    pass # אם נכשל, אולי הסורק הוא Flatbed בלבד
                break
    except Exception:
        pass

def save_wia_image_as_pdf(image_obj, output_path, dpi):
    """פונקציית עזר לשמירת אובייקט WIA כקובץ PDF"""
    temp_bmp_path = None
    pil_image = None
    
    try:
        # שמירה זמנית בטוחה
        temp_dir = tempfile.gettempdir()
        temp_filename = f"scan_{uuid4().hex}.bmp"
        full_temp_path = os.path.join(temp_dir, temp_filename)
        
        # טריק נתיב קצר
        with open(full_temp_path, 'w') as f: pass
        try:
            temp_bmp_path = win32api.GetShortPathName(full_temp_path)
        except:
            temp_bmp_path = full_temp_path
        if os.path.exists(temp_bmp_path): os.remove(temp_bmp_path)

        # שמירת התמונה הפיזית
        image_obj.SaveFile(temp_bmp_path)
        
        # המרה ל-PDF
        pil_image = Image.open(temp_bmp_path)
        if pil_image.mode != "RGB":
            pil_image = pil_image.convert("RGB")
        
        # יצירת תיקיות אם צריך
        output_dir = os.path.dirname(output_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir, exist_ok=True)
            
        pil_image.save(output_path, "PDF", resolution=dpi)
        return True

    except Exception as e:
        print(f"Error saving PDF: {e}")
        return False
    finally:
        if pil_image:
            try: pil_image.close()
            except: pass
        if temp_bmp_path and os.path.exists(temp_bmp_path):
            try: os.remove(temp_bmp_path)
            except: pass

def scan_and_process(output_folder, regex_pattern, log_callback=None):
    """
    סריקת אצווה (Batch Scan):
    סורק את כל הדפים במזין, ולכל דף מבצע OCR, זיהוי ושמירה בתיקייה ייעודית.
    """
    import pdf_processor
    
    pythoncom.CoInitialize()
    
    try:
        if log_callback:
            log_callback("מתחבר לסורק...")
        
        # 1. חיבור לסורק
        device_manager = win32com.client.Dispatch("WIA.DeviceManager")
        if device_manager.DeviceInfos.Count == 0:
            if log_callback: log_callback("✗ לא נמצא סורק מחובר")
            return None
            
        # מתחבר לסורק הראשון (או ניתן להוסיף לוגיקה לבחירה)
        device = device_manager.DeviceInfos(1).Connect()
        
        # 2. ניסיון להגדיר ADF (מזין דפים)
        configure_adf(device)
        
        item = device.Items(1)
        
        # הגדרות סריקה (DPI 300 וצבע)
        for prop in item.Properties:
            if prop.Name == "Horizontal Resolution": prop.Value = 300
            if prop.Name == "Vertical Resolution": prop.Value = 300
            if prop.Name == "Current Intent": prop.Value = 1 # Color

        WIA_FORMAT_BMP = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}"
        
        pages_scanned = 0
        
        if log_callback:
            log_callback("מתחיל סריקת אצווה (כל הדפים במגש)...")

        # === לולאת הסריקה ===
        while True:
            try:
                # ניסיון למשוך דף. אם המגש ריק, WIA יזרוק שגיאה ונצא מהלולאה
                image_obj = item.Transfer(WIA_FORMAT_BMP)
                
                pages_scanned += 1
                if log_callback: log_callback(f"-- מעבד דף מספר {pages_scanned} --")
                
                # יצירת שם זמני לדף הנוכחי
                temp_pdf_name = f"temp_page_{pages_scanned}_{uuid4().hex[:6]}.pdf"
                temp_pdf_path = os.path.join(output_folder, temp_pdf_name)
                
                # שמירה זמנית
                success = save_wia_image_as_pdf(image_obj, temp_pdf_path, 300)
                
                if success:
                    # ביצוע OCR וזיהוי מספר
                    if log_callback: log_callback("   מפענח טקסט (OCR)...")
                    
                    match_value = pdf_processor.process_pdf_file(temp_pdf_path, regex_pattern)
                    
                    if match_value:
                        # שימוש בלוגיקה של תיקיות (מס' ת"ז -> קובץ ממוספר)
                        new_full_path = pdf_processor.generate_id_folder_path(output_folder, match_value)
                        
                        try:
                            os.rename(temp_pdf_path, new_full_path)
                            
                            # לוג יפה למשתמש
                            final_name = os.path.basename(new_full_path)
                            folder_name = os.path.basename(os.path.dirname(new_full_path))
                            if log_callback: log_callback(f"   ✓ זוהה: {match_value} -> נשמר ב: {folder_name}/{final_name}")
                            
                        except OSError as e:
                            if log_callback: log_callback(f"   ✗ שגיאה בהעברה: {e}")
                    else:
                        # אם לא זוהה מספר
                        if log_callback: log_callback("   ⚠ לא זוהה מס' תעודת זהות (נשמר בתיקייה הראשית)")
                        
                else:
                    if log_callback: log_callback("   ✗ שגיאה בשמירת הקובץ הסרוק")

            except Exception as e:
                # שגיאה כאן בדרך כלל אומרת שהמזין ריק (נגמרו הדפים)
                # בדיקה אם זו שגיאה "רגילה" של סיום סריקה
                error_str = str(e)
                if "0x80210003" in error_str or "paper is empty" in error_str.lower():
                    # זה הקוד הרשמי של WIA ל"נגמר הנייר"
                    pass 
                elif pages_scanned > 0:
                    # אם סרקנו כבר דפים ועפנו, כנראה פשוט נגמרו הדפים
                    pass
                else:
                    # אם עפנו על ההתחלה, זו שגיאה אמיתית
                    if log_callback: log_callback(f"סריקה הסתיימה או נעצרה: {e}")
                
                break # יציאה מהלולאה

        if log_callback:
            log_callback(f"\nסיכום: נסרקו ועובדו {pages_scanned} דפים.")
            
        return "Batch Complete"

    except Exception as e:
        if log_callback:
            log_callback(f"שגיאה כללית במודול הסריקה: {e}")
        return None