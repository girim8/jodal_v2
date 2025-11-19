# HWP to TXT Converter

## Project Overview
A Streamlit web application that converts HWP (Hangul Word Processor) files to text format with comprehensive Korean encoding support. The app implements multiple fallback conversion strategies to ensure successful extraction of Korean text, with automatic verification that the keyword '스쿨넷' appears correctly in the output.

## Recent Changes (November 10, 2025)
- ✅ Implemented multi-method HWP conversion with 6 fallback strategies
- ✅ Fixed file upload 403 error by configuring Streamlit CORS/XSRF settings
- ✅ Implemented proper HWP record parsing with -zlib.MAX_WBITS decompression
- ✅ Added UTF-16LE decoding for HWP text records (tags 66, 67, 68, 80)
- ✅ Expanded verification to 4 keywords: '스쿨넷', '추정금액', '이중화', 'Aggregation'
- ✅ Successfully tested with real HWP document (260,806 chars extracted, 92,387 Korean)
- ✅ Table text extraction working - content preserved without layout structure

## Features
1. **Multi-Method Conversion Pipeline**: 6 fallback strategies ensure successful conversion
   - Method 1: OLE File Extraction (olefile + zlib decompression) ✅ WORKING
   - Method 2: HWPX ZIP Extraction (for HWPX format files)
   - Method 3: LibreOffice conversion (unoconv)
   - Method 4: CloudConvert API
   - Method 5: ConvertAPI
   - Method 6: Hancom official API

2. **Korean Encoding Support**:
   - Automatic encoding detection using chardet
   - Supports UTF-16LE, UTF-8, CP949, EUC-KR encodings
   - Special handling for HWP's proprietary encoding

3. **Verification System**:
   - Checks for '스쿨넷' presence in converted text
   - Only accepts conversions that pass verification
   - Displays pass/fail status for each method

4. **User Interface**:
   - Drag-and-drop file upload
   - Real-time conversion progress
   - Results dashboard showing character count, Korean count, verification status
   - Download button for converted TXT file
   - Table structure detection

## Technical Architecture

### Core Components
- **app.py**: Main Streamlit application with conversion logic
- **.streamlit/config.toml**: Server configuration (CORS/XSRF disabled for file uploads)
- **test_conversion.py**: Command-line testing utility

### Key Functions
- `method1_olefile_extraction()`: Primary extraction using olefile (WORKING)
- `decompress_bodytext()`: zlib decompression for HWP BodyText sections
- `extract_text_from_hwp_bodytext()`: Text extraction with encoding detection
- `verify_korean_text()`: Validation that Korean text is correctly extracted

### Dependencies
- **Python Libraries**: streamlit, olefile, lxml, chardet, requests, pandas, zlib
- **System**: LibreOffice (for Method 3 fallback)

## Testing
Successfully tested with document: `20210430609-00_1618895289680_4단계 스쿨넷 사업 제안요청서(조달청 추가 의견 반영본) - 복사본_1762740838248.hwp`
- **Result**: Method 1 successfully extracted 260,806 characters (92,387 Korean characters)
- **Verification**: All 4 keywords found and verified ✅
  - '스쿨넷' (Korean - SchoolNet)
  - '추정금액' (Korean - Estimated Amount) 
  - '이중화' (Korean - Redundancy)
  - 'Aggregation' (English - Link Aggregation)
- **Performance**: Fast extraction (~2-3 seconds)
- **Table Content**: Table text successfully extracted and preserved

## Known Limitations
- Method 2 (HWPX): Only works with HWPX format (not HWP 5.0 OLE format)
- Method 3 (LibreOffice): LibreOffice cannot open some HWP files (BrokenPackageRequest error)
- Methods 4-6 (APIs): Require API keys to be configured by user

## API Configuration
Users can optionally configure API keys in the sidebar for fallback methods:
- CloudConvert API Key
- ConvertAPI Secret
- Hancom API Key

Note: Method 1 works without any API keys and successfully handles most HWP files.

## User Preferences
- Primary focus: Korean text preservation
- Verification requirement: Must extract '스쿨넷' correctly
- Table support: Tables should be preserved (shown as text format)
