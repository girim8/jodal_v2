#!/usr/bin/env python3
"""
HWP Record-based parser for proper text extraction
"""

import struct
import zlib
import olefile

class HWPRecordParser:
    """Parse HWP file records to extract text"""
    
    def __init__(self, ole_file):
        self.ole = ole_file
        
    def read_record_header(self, data, offset):
        """Read HWP record header (4 bytes)"""
        if offset + 4 > len(data):
            return None, None, offset
        
        header = struct.unpack('<I', data[offset:offset+4])[0]
        
        # Extract record type and size from header
        record_id = header & 0x3FF  # Lower 10 bits
        level = (header >> 10) & 0x3FF  # Next 10 bits
        size = (header >> 20) & 0xFFF  # Upper 12 bits
        
        return record_id, size, offset + 4
    
    def decompress_record(self, data):
        """Decompress record data if compressed"""
        try:
            # Try zlib decompression
            return zlib.decompress(data)
        except:
            return data
    
    def extract_text_from_bodytext(self, stream_name='BodyText/Section0'):
        """Extract text from BodyText section"""
        try:
            data = self.ole.openstream(stream_name).read()
        except:
            return ""
        
        extracted_texts = []
        offset = 0
        
        print(f"Parsing {stream_name} ({len(data):,} bytes)...")
        
        while offset < len(data):
            record_id, size, new_offset = self.read_record_header(data, offset)
            
            if record_id is None:
                break
            
            # Read record data
            if new_offset + size > len(data):
                break
            
            record_data = data[new_offset:new_offset + size]
            
            # HWP record types that contain text:
            # 0x50 (HWPTAG_PARA_TEXT) - Paragraph text
            # 0x43 (HWPTAG_PARA_CHAR_SHAPE) - Character formatting
            # 0x45 (HWPTAG_TABLE) - Table data
            
            if record_id == 0x50:  # PARA_TEXT
                # Try to decompress if needed
                try_decompress = self.decompress_record(record_data)
                
                # Try different encodings
                for encoding in ['utf-16le', 'utf-8', 'cp949', 'euc-kr']:
                    try:
                        text = try_decompress.decode(encoding, errors='ignore')
                        # Check if we got meaningful Korean text
                        if any('\uac00' <= c <= '\ud7a3' for c in text):
                            extracted_texts.append(text)
                            break
                    except:
                        continue
            
            offset = new_offset + size
        
        return '\n'.join(extracted_texts)
    
    def extract_all_text(self):
        """Extract text from all available streams"""
        all_text = []
        
        # Get all BodyText sections
        for stream_path in self.ole.listdir():
            stream_name = '/'.join(stream_path)
            
            if 'BodyText/' in stream_name:
                text = self.extract_text_from_bodytext(stream_name)
                if text:
                    all_text.append(text)
        
        return '\n'.join(all_text)


def test_hwp_parser():
    """Test the HWP parser"""
    test_file = 'attached_assets/20210430609-00_1618895289680_4단계 스쿨넷 사업 제안요청서(조달청 추가 의견 반영본) - 복사본_1762740838248.hwp'
    
    KEYWORDS = ['스쿨넷', '추정금액', '이중화', 'Aggregation']
    
    print("="*80)
    print("Testing HWP Record Parser")
    print("="*80)
    
    with open(test_file, 'rb') as f:
        hwp_data = f.read()
    
    ole = olefile.OleFileIO(hwp_data)
    parser = HWPRecordParser(ole)
    
    text = parser.extract_all_text()
    
    print(f"\nExtracted {len(text):,} characters")
    
    korean_count = sum(1 for c in text if '\uac00' <= c <= '\ud7a3')
    print(f"Korean characters: {korean_count:,}")
    
    print("\nKeyword검색:")
    found = []
    for keyword in KEYWORDS:
        if keyword in text:
            print(f"  ✓ {keyword}")
            found.append(keyword)
        else:
            print(f"  ✗ {keyword}")
    
    if len(found) == len(KEYWORDS):
        print("\n✅ SUCCESS! All keywords found!")
    else:
        print(f"\n⚠️ Found {len(found)}/{len(KEYWORDS)} keywords")
    
    print(f"\nSample text (first 500 chars):")
    print("-" * 80)
    print(text[:500])
    
    ole.close()
    
    return text, found


if __name__ == '__main__':
    test_hwp_parser()
