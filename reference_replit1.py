# 아래 코드는 replit으로 작성한, 정상적으로 hwp to txt를 작동시킨 컴팩트한 코드임. 이를 참고하여 코드를 구성할 것 

#!/usr/bin/env python3
"""
Proper HWP parsing based on record structure
"""

import struct
import zlib
import olefile

test_file = 'attached_assets/20210430609-00_1618895289680_4단계 스쿨넷 사업 제안요청서(조달청 추가 의견 반영본) - 복사본_1762740838248.hwp'

KEYWORDS = ['스쿨넷', '추정금액', '이중화', 'Aggregation']

print("="*80)
print("Proper HWP Record Parsing")
print("="*80)

with open(test_file, 'rb') as f:
    hwp_data = f.read()

ole = olefile.OleFileIO(hwp_data)

# Get BodyText/Section0
stream_path = None
for sp in ole.listdir():
    if '/'.join(sp) == 'BodyText/Section0':
        stream_path = sp
        break

if not stream_path:
    print("BodyText/Section0 not found!")
    exit(1)

compressed_data = ole.openstream(stream_path).read()
print(f"\nCompressed size: {len(compressed_data):,} bytes")

# Try decompression with -MAX_WBITS (no zlib header)
print("\nAttempting zlib decompression with -MAX_WBITS...")
try:
    decompressed = zlib.decompress(compressed_data, -zlib.MAX_WBITS)
    print(f"✓ Decompressed successfully: {len(decompressed):,} bytes")
except Exception as e:
    print(f"✗ Failed with -MAX_WBITS: {e}")
    # Try with MAX_WBITS (with header)
    try:
        decompressed = zlib.decompress(compressed_data)
        print(f"✓ Decompressed with header: {len(decompressed):,} bytes")
    except Exception as e2:
        print(f"✗ All decompression failed: {e2}")
        exit(1)

print("\nParsing records...")
print("-" * 80)

# Parse records
offset = 0
texts = []
record_count = 0
para_text_count = 0

while offset < len(decompressed) - 4:
    # Read record header (4 bytes)
    try:
        header = struct.unpack('<I', decompressed[offset:offset+4])[0]
    except:
        break
    
    # Extract fields from header
    tag_id = header & 0x3FF  # bits 0-9
    level = (header >> 10) & 0x3FF  # bits 10-19
    size = (header >> 20) & 0xFFF  # bits 20-31
    
    offset += 4
    record_count += 1
    
    # Check if we have enough data for payload
    if offset + size > len(decompressed):
        break
    
    # Read payload
    payload = decompressed[offset:offset + size]
    offset += size
    
    # HWPTAG_PARA_TEXT is tag 66 (0x42) according to documentation
    # But let's check multiple tags that might contain text
    if tag_id in [66, 67, 68, 80]:  # Try common text-related tags
        para_text_count += 1
        
        # Try to decode as UTF-16LE
        try:
            text = payload.decode('utf-16le', errors='ignore')
            # Filter out control characters, keep Korean and printable
            filtered = ''.join(c for c in text if c.isprintable() or c.isspace())
            
            if filtered and len(filtered) > 2:
                # Check if it contains Korean
                has_korean = any('\uac00' <= c <= '\ud7a3' for c in filtered)
                if has_korean or any(kw in filtered for kw in KEYWORDS):
                    texts.append(filtered)
                    
                    if record_count <= 100:  # Show first few
                        print(f"  Record #{record_count}: tag={tag_id:3d}, level={level}, size={size:5d}")
                        print(f"    Text preview: {filtered[:80]}...")
        except:
            pass

print(f"\nTotal records parsed: {record_count}")
print(f"Para text records: {para_text_count}")
print(f"Text chunks extracted: {len(texts)}")

# Combine all texts
combined_text = '\n'.join(texts)
print(f"\nTotal extracted text length: {len(combined_text):,} characters")

korean_count = sum(1 for c in combined_text if '\uac00' <= c <= '\ud7a3')
print(f"Korean characters: {korean_count:,} ({korean_count/len(combined_text)*100:.1f}%)")

print("\n" + "="*80)
print("Keyword Verification")
print("="*80)

found_keywords = []
for keyword in KEYWORDS:
    if keyword in combined_text:
        print(f"  ✅ Found: {keyword}")
        found_keywords.append(keyword)
        # Show context
        idx = combined_text.find(keyword)
        context = combined_text[max(0, idx-30):idx+len(keyword)+30]
        print(f"     Context: ...{context}...")
    else:
        print(f"  ❌ Missing: {keyword}")

if len(found_keywords) == len(KEYWORDS):
    print("\n🎉 SUCCESS! All keywords found!")
else:
    print(f"\n⚠️ Found {len(found_keywords)}/{len(KEYWORDS)} keywords")

print("\n" + "="*80)
print("Sample Text (first 1000 characters)")
print("="*80)
print(combined_text[:1000])

ole.close()
