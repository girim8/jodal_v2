import streamlit as st
import olefile
import zipfile
import chardet
import io
import os
import zlib
import struct
from lxml import etree
import tempfile
import requests

st.set_page_config(page_title="HWP to TXT Converter", page_icon="📄", layout="wide")

st.title("📄 HWP to TXT Converter")
st.markdown("Convert HWP (Hangul Word Processor) files to text with Korean encoding support")

# Verification keywords
VERIFICATION_KEYWORDS = ["스쿨넷", "추정금액", "이중화", "Aggregation"]
VERIFICATION_KEYWORD = "스쿨넷"  # Primary keyword for backward compatibility

def detect_encoding(data):
    """Detect encoding using chardet"""
    try:
        result = chardet.detect(data)
        return result['encoding'], result['confidence']
    except:
        return 'utf-8', 0.0

def verify_korean_text(text, keyword=VERIFICATION_KEYWORD):
    """Verify if Korean text appears correctly in the converted text"""
    if not text:
        return False, "Empty text"
    
    # Check all verification keywords
    found_keywords = [kw for kw in VERIFICATION_KEYWORDS if kw in text]
    
    if found_keywords:
        keywords_str = ', '.join(found_keywords)
        if len(found_keywords) == len(VERIFICATION_KEYWORDS):
            return True, f"✅ All keywords found: {keywords_str}"
        else:
            missing = set(VERIFICATION_KEYWORDS) - set(found_keywords)
            return True, f"✅ Found {len(found_keywords)}/{len(VERIFICATION_KEYWORDS)}: {keywords_str} (missing: {', '.join(missing)})"
    else:
        # Check if text contains any Korean characters
        has_korean = any('\uac00' <= char <= '\ud7a3' for char in text)
        if has_korean:
            return False, f"⚠️ Korean text found but keywords missing: {', '.join(VERIFICATION_KEYWORDS)}"
        else:
            return False, f"❌ No Korean text detected ('{keyword}' not found)"

def parse_hwp_records(data):
    """Parse HWP records from decompressed BodyText section"""
    texts = []
    offset = 0
    
    while offset < len(data) - 4:
        try:
            # Read record header (4 bytes)
            header = struct.unpack('<I', data[offset:offset+4])[0]
            
            # Extract fields from header
            tag_id = header & 0x3FF  # bits 0-9
            level = (header >> 10) & 0x3FF  # bits 10-19  
            size = (header >> 20) & 0xFFF  # bits 20-31
            
            offset += 4
            
            # Check if we have enough data for payload
            if offset + size > len(data):
                break
            
            # Read payload
            payload = data[offset:offset + size]
            offset += size
            
            # Tag 67 typically contains paragraph text
            # Also check tags 66, 68, 80 for text content
            if tag_id in [66, 67, 68, 80]:
                try:
                    # Decode as UTF-16LE (standard for HWP)
                    text = payload.decode('utf-16le', errors='ignore')
                    # Filter out control characters, keep printable
                    filtered = ''.join(c for c in text if c.isprintable() or c.isspace())
                    
                    if filtered and len(filtered) > 2:
                        # Check if it contains Korean or meaningful text
                        has_korean = any('\uac00' <= c <= '\ud7a3' for c in filtered)
                        has_english = any(c.isascii() and c.isalpha() for c in filtered)
                        
                        if has_korean or has_english:
                            texts.append(filtered)
                except:
                    pass
        except:
            break
    
    return '\n'.join(texts)

def extract_text_from_hwp_bodytext(compressed_data):
    """Extract text from HWP BodyText section with proper decompression"""
    try:
        # HWP BodyText uses zlib compression without header
        # Use -MAX_WBITS flag for raw DEFLATE
        try:
            decompressed = zlib.decompress(compressed_data, -zlib.MAX_WBITS)
        except:
            # Fallback: try with header
            try:
                decompressed = zlib.decompress(compressed_data)
            except:
                # Not compressed, use as-is
                decompressed = compressed_data
        
        # Parse HWP records to extract text
        text = parse_hwp_records(decompressed)
        
        if text and len(text) > 100:
            return text
        else:
            return None
            
    except Exception as e:
        return None

def method1_olefile_extraction(hwp_data):
    """
    Method 1: Extract text from HWP using olefile (for HWP 5.0+)
    """
    try:
        st.write("🔍 **Method 1:** Trying OLE File extraction...")
        
        # Create a file-like object from bytes
        ole = olefile.OleFileIO(hwp_data)
        
        # List all streams in the OLE file
        streams = ole.listdir()
        st.write(f"  Found {len(streams)} streams in OLE file")
        
        extracted_text = []
        bodytext_found = False
        
        for stream_path in streams:
            stream_name = '/'.join(stream_path)
            
            try:
                data = ole.openstream(stream_path).read()
                
                # Special handling for BodyText sections (main content)
                if 'BodyText' in stream_name:
                    bodytext_found = True
                    st.write(f"  Processing {stream_name} ({len(data):,} bytes)...")
                    text = extract_text_from_hwp_bodytext(data)
                    if text and len(text) > 100:
                        # Check for Korean content
                        korean_chars = sum(1 for c in text if '\uac00' <= c <= '\ud7a3')
                        if korean_chars > 50:
                            st.write(f"    Found {korean_chars} Korean characters")
                            extracted_text.append(text)
                else:
                    # Try regular decoding for other streams
                    for encoding in ['utf-16le', 'utf-8', 'cp949', 'euc-kr']:
                        try:
                            text = data.decode(encoding, errors='ignore')
                            if text and len(text) > 10:
                                # Check if this stream has meaningful Korean text
                                korean_chars = sum(1 for c in text if '\uac00' <= c <= '\ud7a3')
                                if korean_chars > 5:
                                    extracted_text.append(text)
                                break
                        except:
                            continue
            except Exception as e:
                continue
        
        ole.close()
        
        if extracted_text:
            combined_text = '\n'.join(extracted_text)
            # Clean up the text - keep Korean, alphanumeric, and basic punctuation
            cleaned = ''.join(char for char in combined_text if char.isprintable() or char.isspace())
            
            # Count statistics
            total_chars = len(cleaned)
            korean_chars = sum(1 for c in cleaned if '\uac00' <= c <= '\ud7a3')
            
            st.write(f"  Extracted {total_chars:,} characters ({korean_chars:,} Korean)")
            is_valid, msg = verify_korean_text(cleaned)
            st.write(f"  {msg}")
            
            if is_valid:
                return True, cleaned, "OLE File Extraction"
            else:
                return False, cleaned, f"OLE File Extraction failed verification: {msg}"
        else:
            st.write("  ❌ No text streams found")
            return False, None, "No text streams in OLE file"
            
    except Exception as e:
        st.write(f"  ❌ Error: {str(e)}")
        return False, None, f"OLE extraction error: {str(e)}"

def method2_hwpx_zipfile_extraction(hwp_data):
    """
    Method 2: Try to extract as HWPX (ZIP format) using zipfile
    """
    try:
        st.write("🔍 **Method 2:** Trying HWPX (ZIP) extraction...")
        
        # Create a file-like object
        zip_file = zipfile.ZipFile(io.BytesIO(hwp_data))
        file_list = zip_file.namelist()
        
        st.write(f"  Found {len(file_list)} files in HWPX archive")
        
        extracted_text = []
        
        # HWPX files contain XML files with content
        for filename in file_list:
            if filename.endswith('.xml'):
                try:
                    xml_data = zip_file.read(filename)
                    
                    # Try to parse as XML
                    try:
                        root = etree.fromstring(xml_data)
                        
                        # Extract all text from XML
                        text_elements = root.xpath('//text()')
                        text = ' '.join(text_elements)
                        
                        if text and len(text) > 10:
                            extracted_text.append(text)
                    except:
                        # If XML parsing fails, try as plain text
                        for encoding in ['utf-8', 'utf-16le', 'cp949', 'euc-kr']:
                            try:
                                text = xml_data.decode(encoding, errors='ignore')
                                if text and len(text) > 10:
                                    extracted_text.append(text)
                                    break
                            except:
                                continue
                except:
                    continue
        
        zip_file.close()
        
        if extracted_text:
            combined_text = '\n'.join(extracted_text)
            cleaned = ''.join(char for char in combined_text if char.isprintable() or char.isspace())
            
            st.write(f"  Extracted {len(cleaned)} characters")
            is_valid, msg = verify_korean_text(cleaned)
            st.write(f"  {msg}")
            
            if is_valid:
                return True, cleaned, "HWPX ZIP Extraction"
            else:
                return False, cleaned, f"HWPX extraction failed verification: {msg}"
        else:
            st.write("  ❌ No text content found in XML files")
            return False, None, "No text in HWPX XML files"
            
    except zipfile.BadZipFile:
        st.write("  ❌ Not a valid ZIP/HWPX file")
        return False, None, "Not a HWPX file"
    except Exception as e:
        st.write(f"  ❌ Error: {str(e)}")
        return False, None, f"HWPX extraction error: {str(e)}"

def method3_libreoffice_conversion(hwp_data):
    """
    Method 3: Convert using LibreOffice (unoconv)
    """
    try:
        st.write("🔍 **Method 3:** Trying LibreOffice conversion...")
        
        # Save HWP to temp file
        with tempfile.NamedTemporaryFile(delete=False, suffix='.hwp') as tmp_hwp:
            tmp_hwp.write(hwp_data)
            tmp_hwp_path = tmp_hwp.name
        
        # Try unoconv conversion
        import subprocess
        
        output_path = tmp_hwp_path.replace('.hwp', '.txt')
        
        # Try unoconv
        result = subprocess.run(
            ['unoconv', '-f', 'txt', '-o', output_path, tmp_hwp_path],
            capture_output=True,
            timeout=30
        )
        
        if os.path.exists(output_path):
            with open(output_path, 'rb') as f:
                text_data = f.read()
            
            # Detect encoding
            encoding, confidence = detect_encoding(text_data)
            st.write(f"  Detected encoding: {encoding} (confidence: {confidence:.2f})")
            
            text = text_data.decode(encoding if encoding else 'utf-8', errors='ignore')
            
            # Clean up temp files
            os.unlink(tmp_hwp_path)
            os.unlink(output_path)
            
            st.write(f"  Extracted {len(text)} characters")
            is_valid, msg = verify_korean_text(text)
            st.write(f"  {msg}")
            
            if is_valid:
                return True, text, "LibreOffice Conversion"
            else:
                return False, text, f"LibreOffice failed verification: {msg}"
        else:
            os.unlink(tmp_hwp_path)
            st.write("  ❌ Conversion produced no output")
            return False, None, "LibreOffice conversion failed"
            
    except FileNotFoundError:
        st.write("  ❌ LibreOffice/unoconv not installed")
        return False, None, "LibreOffice not available"
    except Exception as e:
        st.write(f"  ❌ Error: {str(e)}")
        return False, None, f"LibreOffice error: {str(e)}"

def method4_cloudconvert_api(hwp_data, api_key):
    """
    Method 4: Convert using CloudConvert API
    """
    try:
        st.write("🔍 **Method 4:** Trying CloudConvert API...")
        
        if not api_key:
            st.write("  ⚠️ API key not provided")
            return False, None, "CloudConvert API key missing"
        
        # CloudConvert API endpoint
        url = "https://api.cloudconvert.com/v2/jobs"
        
        headers = {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json"
        }
        
        # Create conversion job
        job_data = {
            "tasks": {
                "import-hwp": {
                    "operation": "import/upload"
                },
                "convert-to-txt": {
                    "operation": "convert",
                    "input": "import-hwp",
                    "output_format": "txt"
                },
                "export-txt": {
                    "operation": "export/url",
                    "input": "convert-to-txt"
                }
            }
        }
        
        response = requests.post(url, headers=headers, json=job_data, timeout=30)
        
        if response.status_code == 201:
            job = response.json()
            
            # Upload file
            upload_task = job['data']['tasks'][0]
            upload_url = upload_task['result']['form']['url']
            
            files = {'file': ('document.hwp', hwp_data, 'application/x-hwp')}
            upload_response = requests.post(upload_url, files=files, timeout=60)
            
            if upload_response.status_code == 200:
                # Wait for conversion (simplified - should poll status)
                import time
                time.sleep(10)
                
                # Get result
                job_id = job['data']['id']
                status_response = requests.get(f"{url}/{job_id}", headers=headers, timeout=30)
                
                if status_response.status_code == 200:
                    result = status_response.json()
                    export_task = [t for t in result['data']['tasks'] if t['operation'] == 'export/url'][0]
                    
                    if export_task['status'] == 'finished':
                        download_url = export_task['result']['files'][0]['url']
                        text_response = requests.get(download_url, timeout=30)
                        
                        if text_response.status_code == 200:
                            text_data = text_response.content
                            encoding, confidence = detect_encoding(text_data)
                            text = text_data.decode(encoding if encoding else 'utf-8', errors='ignore')
                            
                            st.write(f"  Extracted {len(text)} characters")
                            is_valid, msg = verify_korean_text(text)
                            st.write(f"  {msg}")
                            
                            if is_valid:
                                return True, text, "CloudConvert API"
                            else:
                                return False, text, f"CloudConvert failed verification: {msg}"
        
        st.write(f"  ❌ API request failed: {response.status_code}")
        return False, None, f"CloudConvert API error: {response.status_code}"
        
    except Exception as e:
        st.write(f"  ❌ Error: {str(e)}")
        return False, None, f"CloudConvert error: {str(e)}"

def method5_convertapi(hwp_data, api_secret):
    """
    Method 5: Convert using ConvertAPI
    """
    try:
        st.write("🔍 **Method 5:** Trying ConvertAPI...")
        
        if not api_secret:
            st.write("  ⚠️ API secret not provided")
            return False, None, "ConvertAPI secret missing"
        
        url = f"https://v2.convertapi.com/convert/hwp/to/txt?Secret={api_secret}"
        
        files = {'File': ('document.hwp', hwp_data, 'application/x-hwp')}
        
        response = requests.post(url, files=files, timeout=60)
        
        if response.status_code == 200:
            result = response.json()
            
            if 'Files' in result and len(result['Files']) > 0:
                file_url = result['Files'][0]['Url']
                text_response = requests.get(file_url, timeout=30)
                
                if text_response.status_code == 200:
                    text_data = text_response.content
                    encoding, confidence = detect_encoding(text_data)
                    text = text_data.decode(encoding if encoding else 'utf-8', errors='ignore')
                    
                    st.write(f"  Extracted {len(text)} characters")
                    is_valid, msg = verify_korean_text(text)
                    st.write(f"  {msg}")
                    
                    if is_valid:
                        return True, text, "ConvertAPI"
                    else:
                        return False, text, f"ConvertAPI failed verification: {msg}"
        
        st.write(f"  ❌ API request failed: {response.status_code}")
        return False, None, f"ConvertAPI error: {response.status_code}"
        
    except Exception as e:
        st.write(f"  ❌ Error: {str(e)}")
        return False, None, f"ConvertAPI error: {str(e)}"

def method6_hancom_api(hwp_data, api_credentials):
    """
    Method 6: Convert using Hancom official API
    """
    try:
        st.write("🔍 **Method 6:** Trying Hancom API...")
        
        if not api_credentials or 'api_key' not in api_credentials:
            st.write("  ⚠️ API credentials not provided")
            return False, None, "Hancom API credentials missing"
        
        # Hancom API endpoint (placeholder - actual endpoint may vary)
        url = "https://api.hancom.com/convert/hwp-to-text"
        
        headers = {
            "Authorization": f"Bearer {api_credentials['api_key']}",
            "Content-Type": "application/octet-stream"
        }
        
        response = requests.post(url, headers=headers, data=hwp_data, timeout=60)
        
        if response.status_code == 200:
            text_data = response.content
            encoding, confidence = detect_encoding(text_data)
            text = text_data.decode(encoding if encoding else 'utf-8', errors='ignore')
            
            st.write(f"  Extracted {len(text)} characters")
            is_valid, msg = verify_korean_text(text)
            st.write(f"  {msg}")
            
            if is_valid:
                return True, text, "Hancom API"
            else:
                return False, text, f"Hancom API failed verification: {msg}"
        else:
            st.write(f"  ❌ API request failed: {response.status_code}")
            return False, None, f"Hancom API error: {response.status_code}"
        
    except Exception as e:
        st.write(f"  ❌ Error: {str(e)}")
        return False, None, f"Hancom API error: {str(e)}"

# Sidebar for API credentials
st.sidebar.header("⚙️ API Configuration")
st.sidebar.markdown("Optional: Provide API keys for fallback conversion methods")

cloudconvert_key = st.sidebar.text_input("CloudConvert API Key", type="password", help="Get from cloudconvert.com")
convertapi_secret = st.sidebar.text_input("ConvertAPI Secret", type="password", help="Get from convertapi.com")
hancom_api_key = st.sidebar.text_input("Hancom API Key", type="password", help="Official Hancom API key")

# Main interface
uploaded_file = st.file_uploader("Upload HWP file", type=['hwp'], help="Upload a Hangul Word Processor file")

if uploaded_file is not None:
    st.success(f"File uploaded: {uploaded_file.name} ({uploaded_file.size:,} bytes)")
    
    # Read file data
    hwp_data = uploaded_file.read()
    
    st.markdown("---")
    st.subheader("🔄 Conversion Process")
    st.info(f"**Verification:** Checking for Korean text '{VERIFICATION_KEYWORD}' in output")
    
    conversion_results = []
    successful_conversion = None
    
    # Try each method in sequence
    methods = [
        ("Method 1: OLE File Extraction", lambda: method1_olefile_extraction(hwp_data)),
        ("Method 2: HWPX ZIP Extraction", lambda: method2_hwpx_zipfile_extraction(hwp_data)),
        ("Method 3: LibreOffice", lambda: method3_libreoffice_conversion(hwp_data)),
        ("Method 4: CloudConvert API", lambda: method4_cloudconvert_api(hwp_data, cloudconvert_key)),
        ("Method 5: ConvertAPI", lambda: method5_convertapi(hwp_data, convertapi_secret)),
        ("Method 6: Hancom API", lambda: method6_hancom_api(hwp_data, {'api_key': hancom_api_key})),
    ]
    
    for method_name, method_func in methods:
        with st.expander(method_name, expanded=True):
            success, text, message = method_func()
            
            conversion_results.append({
                'method': method_name,
                'success': success,
                'text': text,
                'message': message
            })
            
            if success:
                successful_conversion = {'method': method_name, 'text': text, 'message': message}
                st.success(f"✅ **SUCCESS!** {method_name} worked!")
                break
    
    st.markdown("---")
    st.subheader("📊 Conversion Results")
    
    # Display results summary
    result_data = []
    for result in conversion_results:
        status = "✅ PASS" if result['success'] else "❌ FAIL"
        result_data.append({
            'Method': result['method'],
            'Status': status,
            'Message': result['message']
        })
    
    import pandas as pd
    df = pd.DataFrame(result_data)
    st.dataframe(df, use_container_width=True)
    
    if successful_conversion:
        st.markdown("---")
        st.subheader("✨ Converted Text")
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Characters", len(successful_conversion['text']))
        with col2:
            korean_chars = sum(1 for c in successful_conversion['text'] if '\uac00' <= c <= '\ud7a3')
            st.metric("Korean Characters", korean_chars)
        with col3:
            st.metric("Verification", "✅ PASSED" if VERIFICATION_KEYWORD in successful_conversion['text'] else "❌ FAILED")
        
        # Display text
        st.text_area("Converted Text", successful_conversion['text'], height=400)
        
        # Download button
        st.download_button(
            label="📥 Download TXT",
            data=successful_conversion['text'].encode('utf-8'),
            file_name=uploaded_file.name.replace('.hwp', '.txt'),
            mime='text/plain'
        )
        
        # Check for tables (simple heuristic)
        if '\t' in successful_conversion['text'] or '|' in successful_conversion['text']:
            st.info("📋 Tables detected in the document (preserved in text format)")
    else:
        st.error("❌ All conversion methods failed. Please check:")
        st.markdown("""
        - The HWP file is not corrupted
        - API keys are correctly configured (if using API methods)
        - LibreOffice is installed on the system (for Method 3)
        """)
else:
    st.info("👆 Upload an HWP file to begin conversion")
    
    # Display test file info
    st.markdown("---")
    st.subheader("📝 About This Converter")
    st.markdown("""
    This converter uses multiple fallback methods to extract text from HWP files:
    
    1. **OLE File Extraction** - Direct parsing of HWP 5.0+ format
    2. **HWPX ZIP Extraction** - Extract from HWPX (XML-based) format
    3. **LibreOffice Conversion** - Using unoconv with LibreOffice
    4. **CloudConvert API** - Cloud-based conversion service
    5. **ConvertAPI** - Alternative cloud conversion service
    6. **Hancom API** - Official Hancom conversion API
    
    Each method is verified to ensure Korean text ('{0}') is correctly extracted.
    """.format(VERIFICATION_KEYWORD))
