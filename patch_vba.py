"""
Patch vbaProject.bin inside test_260525_03.xlsm:
1. ParseSupportReaction: remove 'Option Private Module'
2. Module1: clear to stub
3. UserForm1: clear to stub
4. UserForm2: rename VB_Name to FormNodeSelect
5. Module01: add public wrapper subs
6. dir stream: rename UserForm2 -> FormNodeSelect in MODULENAME records
7. PROJECT stream: update BaseClass= and workspace entry
8. OLE directory entries: rename 'UserForm2' -> 'FormNodeSelect'
"""

import struct, math, zipfile, io, re, sys
import olefile
from oletools.olevba import decompress_stream

XLSM_PATH = '/home/user/260526_01/test_260525_03.xlsm'


# ─── MS-OVBA Compression ────────────────────────────────────────────────────

def _copytoken_help(local_pos: int):
    difference = max(local_pos, 1)
    bit_count = max(int(math.ceil(math.log(difference, 2))), 4)
    length_mask = 0xFFFF >> bit_count
    maximum_length = length_mask + 3
    return bit_count, length_mask, maximum_length


def _compress_chunk(chunk: bytes) -> bytes:
    """Return compressed bytes for one chunk (up to 4096 bytes), including 2-byte header."""
    data = bytearray(chunk)
    n = len(data)
    tokens = bytearray()
    pos = 0
    match_map: dict = {}

    while pos < n:
        flag_idx = len(tokens)
        tokens.append(0)
        flag = 0

        for bit in range(8):
            if pos >= n:
                break
            bit_count, lm, max_len = _copytoken_help(pos)
            best_offset, best_length = 0, 0

            if pos >= 3 and pos + 3 <= n:
                key = (data[pos], data[pos + 1], data[pos + 2])
                for prev in reversed(match_map.get(key, [])):
                    offset = pos - prev
                    if offset > 4096:
                        break
                    length = 0
                    while length < max_len and pos + length < n:
                        if data[prev + (length % offset)] != data[pos + length]:
                            break
                        length += 1
                    if length > best_length:
                        best_length, best_offset = length, offset
                    if best_length >= max_len:
                        break

            # Update hash map
            if pos + 3 <= n:
                k = (data[pos], data[pos + 1], data[pos + 2])
                match_map.setdefault(k, []).append(pos)

            if best_length >= 3:
                flag |= (1 << bit)
                bit_count, lm, max_len = _copytoken_help(pos)
                best_length = min(best_length, max_len)
                temp2 = 16 - bit_count
                copy_token = ((best_offset - 1) << temp2) | (best_length - 3)
                tokens.extend(struct.pack('<H', copy_token))
                for k in range(1, best_length):
                    p = pos + k
                    if p + 3 <= n:
                        key2 = (data[p], data[p + 1], data[p + 2])
                        match_map.setdefault(key2, []).append(p)
                pos += best_length
            else:
                tokens.append(data[pos])
                pos += 1

        tokens[flag_idx] = flag

    n_comp = len(tokens)
    # If compressed ≥ raw size, use raw chunk (only for full 4096-byte chunks)
    if n == 4096 and n_comp + 2 >= 4098:
        hdr = bytearray(struct.pack('<H', 0x3FFF))
        hdr.extend(data)
        return bytes(hdr)
    # MS-OVBA: header low 12 bits = (data_bytes - 1), total chunk = data_bytes + 2
    # Minimum data is 1 byte per spec, but keep n_comp >= 1 for safety
    if n_comp < 1:
        tokens.append(0)
        n_comp = 1
    header = (n_comp - 1) | 0xB000
    result = bytearray(struct.pack('<H', header))
    result.extend(tokens)
    return bytes(result)


def compress_stream(data: bytes) -> bytes:
    result = bytearray([0x01])
    pos = 0
    while pos < len(data):
        result.extend(_compress_chunk(data[pos:pos + 4096]))
        pos += 4096
    return bytes(result)


# ─── OLE helpers ─────────────────────────────────────────────────────────────

def find_compressed_offset(stream_data: bytes) -> int:
    """Find the byte offset in stream_data where MS-OVBA compressed source begins."""
    for i in range(len(stream_data) - 2):
        if stream_data[i] != 0x01:
            continue
        header = struct.unpack_from('<H', stream_data, i + 1)[0]
        chunk_sig = (header >> 12) & 0x07
        chunk_flag = (header >> 15) & 0x01
        chunk_size = (header & 0x0FFF) + 3
        if chunk_sig != 0b011:
            continue
        if not ((chunk_flag == 1 and 3 <= chunk_size <= 4098) or
                (chunk_flag == 0 and chunk_size == 4098)):
            continue
        try:
            result = decompress_stream(bytearray(stream_data[i:]))
            if len(result) > 10 and b'Attribute VB_Name' in result[:200]:
                return i
        except Exception:
            pass
    return -1


def get_stream(ole, path: str) -> bytes:
    return ole.openstream(path).read()


# ─── dir stream renamer ───────────────────────────────────────────────────────

def rename_in_dir_stream(dir_compressed: bytes, old_name: str, new_name: str) -> bytes:
    """
    Decompresses the dir stream, renames old_name -> new_name in all
    MODULENAME (0x0019), MODULENAMEUNICODE (0x0047),
    MODULESTREAMNAME (0x001A), MODULESTREAMNAMEUNICODE (0x0032) records,
    then recompresses.
    """
    dec = bytearray(decompress_stream(bytearray(dir_compressed)))
    old_ansi = old_name.encode('latin-1')
    new_ansi = new_name.encode('latin-1')
    old_utf16 = old_name.encode('utf-16-le')
    new_utf16 = new_name.encode('utf-16-le')

    result = bytearray()
    s = io.BytesIO(dec)

    while True:
        rid_bytes = s.read(2)
        if len(rid_bytes) < 2:
            break
        rid = struct.unpack('<H', rid_bytes)[0]

        # PROJECTVERSION has special format (no standard size field)
        if rid == 0x0009:
            reserved = s.read(4)
            major = s.read(4)
            minor = s.read(2)
            result.extend(rid_bytes + reserved + major + minor)
            continue

        sz_bytes = s.read(4)
        if len(sz_bytes) < 4:
            result.extend(rid_bytes + sz_bytes)
            break
        sz = struct.unpack('<I', sz_bytes)[0]
        data = s.read(sz)

        if rid in (0x0019, 0x001A):  # MODULENAME, MODULESTREAMNAME (ANSI)
            if data == old_ansi:
                data = new_ansi
                sz = len(data)
                sz_bytes = struct.pack('<I', sz)
        elif rid in (0x0047, 0x0032):  # MODULENAMEUNICODE, MODULESTREAMNAMEUNICODE
            if data == old_utf16:
                data = new_utf16
                sz = len(data)
                sz_bytes = struct.pack('<I', sz)

        result.extend(rid_bytes + sz_bytes + data)

    return compress_stream(bytes(result))


# ─── OLE binary patcher ───────────────────────────────────────────────────────

def patch_ole_stream(vba_bin: bytearray, ole, stream_path: str,
                     new_stream_bytes: bytes) -> bool:
    """
    Replace a stream's full content in the OLE binary (in-place).
    Updates the OLE directory entry's size field to match new content.
    Returns True on success.
    """
    path_parts = stream_path.split('/')
    sector_size = 1 << struct.unpack_from('<H', vba_bin, 30)[0]
    mini_sector_size = 1 << struct.unpack_from('<H', vba_bin, 32)[0]

    # ── locate FAT ──
    num_fat = struct.unpack_from('<I', vba_bin, 44)[0]
    fat_secs = []
    for i in range(min(num_fat, 109)):
        s = struct.unpack_from('<I', vba_bin, 76 + i * 4)[0]
        if s != 0xFFFFFFFF:
            fat_secs.append(s)

    def fat(sec):
        entries = sector_size // 4
        for i, fs in enumerate(fat_secs):
            base = i * entries
            if sec < base + entries:
                off = 512 + fs * sector_size + (sec - base) * 4
                return struct.unpack_from('<I', vba_bin, off)[0]
        return 0xFFFFFFFE

    def sec_off(sec):
        return 512 + sec * sector_size

    # ── locate mini-FAT ──
    minifat_first = struct.unpack_from('<I', vba_bin, 60)[0]

    def minifat(ms):
        entries = sector_size // 4
        sec = minifat_first
        idx = ms // entries
        for _ in range(idx):
            sec = fat(sec)
        off = sec_off(sec) + (ms % entries) * 4
        return struct.unpack_from('<I', vba_bin, off)[0]

    # Root entry's start sector (for mini-stream)
    root_dir_sec = struct.unpack_from('<I', vba_bin, 48)[0]
    root_entry_off = sec_off(root_dir_sec)
    root_start = struct.unpack_from('<I', vba_bin, root_entry_off + 116)[0]

    def mini_sec_off(ms):
        """File offset of mini-sector ms."""
        mss = sector_size // mini_sector_size
        main_sec_idx = ms // mss
        sec = root_start
        for _ in range(main_sec_idx):
            sec = fat(sec)
        return sec_off(sec) + (ms % mss) * mini_sector_size

    # ── find target entry ──
    entry_idx = ole._find(path_parts)
    if entry_idx is None:
        print(f"  ERROR: entry not found for {stream_path}")
        return False
    entry = ole.direntries[entry_idx]

    start_sec = entry.isectStart
    orig_size = entry.size
    use_mini = entry.is_minifat

    new_logical_size = len(new_stream_bytes)  # actual new content size (before padding)

    if new_logical_size > orig_size:
        print(f"  WARNING: {stream_path}: new size {new_logical_size} > orig {orig_size}. Truncating.")
        new_stream_bytes = new_stream_bytes[:orig_size]
        new_logical_size = orig_size

    # Pad to fill the same sector allocation (avoids leaving stale bytes in last sector)
    padded = new_stream_bytes + b'\x00' * (orig_size - new_logical_size)

    # ── write sector chain ──
    written = 0
    sec = start_sec
    chunk = mini_sector_size if use_mini else sector_size
    while sec < 0xFFFFFFFE and written < len(padded):
        if use_mini:
            file_off = mini_sec_off(sec)
            next_sec = minifat(sec)
        else:
            file_off = sec_off(sec)
            next_sec = fat(sec)
        to_write = padded[written:written + chunk]
        vba_bin[file_off:file_off + len(to_write)] = to_write
        written += chunk
        sec = next_sec

    # ── Update OLE directory entry size so olefile reads only new_logical_size bytes ──
    # Without this, olefile reads orig_size bytes including trailing zeros,
    # causing decompressors to fail when they hit the zero padding.
    entries_per_sec = sector_size // 128
    dir_chain_pos = entry_idx // entries_per_sec
    within_sec_pos = entry_idx % entries_per_sec
    dir_chain_sec = root_dir_sec
    for _ in range(dir_chain_pos):
        dir_chain_sec = fat(dir_chain_sec)
    dir_entry_off = sec_off(dir_chain_sec) + within_sec_pos * 128
    struct.pack_into('<I', vba_bin, dir_entry_off + 120, new_logical_size)

    return True


def patch_ole_dir_entries(vba_bin: bytearray, old_name: str, new_name: str):
    """Rename all OLE directory entries named old_name to new_name."""
    sector_size = 1 << struct.unpack_from('<H', vba_bin, 30)[0]
    first_dir_sec = struct.unpack_from('<I', vba_bin, 48)[0]

    old_utf16 = (old_name + '\x00').encode('utf-16-le')
    new_utf16_padded = bytearray(64)
    new_encoded = (new_name + '\x00').encode('utf-16-le')
    new_utf16_padded[:len(new_encoded)] = new_encoded
    new_name_len = len(new_encoded)

    def fat(sec):
        entries = sector_size // 4
        # Assume single FAT sector at sector 0
        off = 512 + struct.unpack_from('<I', vba_bin, 76)[0] * sector_size + sec * 4
        return struct.unpack_from('<I', vba_bin, off)[0]

    sec = first_dir_sec
    visited = set()
    while sec < 0xFFFFFFF8 and sec not in visited:
        visited.add(sec)
        sec_start = 512 + sec * sector_size
        for i in range(sector_size // 128):
            entry_off = sec_start + i * 128
            if entry_off + 128 > len(vba_bin):
                break
            obj_type = vba_bin[entry_off + 66]
            if obj_type == 0:
                continue
            name_len = struct.unpack_from('<H', vba_bin, entry_off + 64)[0]
            if name_len == len(old_utf16):
                name_bytes = bytes(vba_bin[entry_off:entry_off + name_len])
                if name_bytes == old_utf16:
                    vba_bin[entry_off:entry_off + 64] = new_utf16_padded
                    struct.pack_into('<H', vba_bin, entry_off + 64, new_name_len)
                    print(f"  OLE dir: renamed {old_name!r} -> {new_name!r} at entry offset {entry_off}")
        next_sec = fat(sec)
        if next_sec >= 0xFFFFFFF8:
            break
        sec = next_sec


# ─── Main ─────────────────────────────────────────────────────────────────────

def main():
    print(f"Reading {XLSM_PATH}")
    with open(XLSM_PATH, 'rb') as f:
        xlsm_orig = f.read()

    zin = zipfile.ZipFile(io.BytesIO(xlsm_orig), 'r')
    vba_bin_orig = zin.read('xl/vbaProject.bin')
    vba_bin = bytearray(vba_bin_orig)
    ole = olefile.OleFileIO(io.BytesIO(vba_bin_orig))

    # ── Collect module stream info ──
    targets = ['VBA/Module1', 'VBA/ParseSupportReaction',
               'VBA/UserForm1', 'VBA/UserForm2', 'VBA/Module01']
    info = {}
    for path in targets:
        raw = get_stream(ole, path)
        offset = find_compressed_offset(raw)
        if offset < 0:
            print(f"  WARN: cannot find compressed offset for {path}")
            continue
        src = decompress_stream(bytearray(raw[offset:]))
        info[path] = {
            'raw': raw,
            'offset': offset,
            'source': src.decode('cp932', errors='replace'),
            'orig_comp_size': len(raw) - offset,
        }
        print(f"  {path}: offset={offset}, src={len(src)}b, orig_comp={len(raw)-offset}b")

    # ── Build new sources ──
    new_sources = {}

    # ParseSupportReaction: remove Option Private Module
    psr = info['VBA/ParseSupportReaction']['source']
    psr_new = re.sub(r'Option Private Module\r?\n', '', psr, count=1)
    new_sources['VBA/ParseSupportReaction'] = psr_new
    print(f"\nParseSupportReaction: -{len(psr)-len(psr_new)} chars")

    # Module1: stub
    new_sources['VBA/Module1'] = 'Attribute VB_Name = "Module1"\r\n'

    # UserForm1: stub
    new_sources['VBA/UserForm1'] = (
        'Attribute VB_Name = "UserForm1"\r\n'
        'Attribute VB_GlobalNameSpace = False\r\n'
        'Attribute VB_Creatable = False\r\n'
        'Attribute VB_PredeclaredId = True\r\n'
        'Attribute VB_Exposed = False\r\n'
    )

    # UserForm2: rename VB_Name
    uf2 = info['VBA/UserForm2']['source']
    new_sources['VBA/UserForm2'] = uf2.replace(
        'Attribute VB_Name = "UserForm2"',
        'Attribute VB_Name = "FormNodeSelect"', 1)

    # Module01: add wrapper subs
    m01 = info['VBA/Module01']['source']
    wrappers = (
        '\r\nPublic Sub RunParseAndSelectNodes()\r\n'
        '    ParseSupportReaction.ParseAndSelectNodes\r\n'
        'End Sub\r\n'
        '\r\nPublic Sub RunFilterByNode()\r\n'
        '    ParseSupportReaction.FilterByNode\r\n'
        'End Sub\r\n'
    )
    new_sources['VBA/Module01'] = m01.rstrip('\r\n') + '\r\n' + wrappers

    # ── Compress and patch each module stream ──
    print("\nCompressing and patching module streams...")
    for path, new_src in new_sources.items():
        i = info[path]
        enc = new_src.encode('cp932', errors='replace')
        comp = compress_stream(enc)
        orig_comp = i['orig_comp_size']
        print(f"  {path}: new_comp={len(comp)}, orig_comp={orig_comp} "
              f"({'OK' if len(comp) <= orig_comp else 'OVERFLOW'})")

        # Build new full stream: pcode prefix + compressed source (no zero padding;
        # patch_ole_stream handles padding and updates the directory entry size)
        pcode = i['raw'][:i['offset']]
        if len(comp) <= orig_comp:
            new_raw = pcode + comp
        else:
            # Overflow: use raw uncompressed chunks (always ≤ source size * 1.0025)
            print(f"    Overflow! Using raw chunks (no LZ). Re-encoding...")
            comp2 = bytearray([0x01])
            pos2 = 0
            while pos2 < len(enc):
                chunk_data = bytearray(enc[pos2:pos2 + 4096])
                if len(chunk_data) < 4096:
                    chunk_data.extend(b'\x00' * (4096 - len(chunk_data)))
                comp2.extend(struct.pack('<H', 0x3FFF))
                comp2.extend(chunk_data)
                pos2 += 4096
            comp = bytes(comp2)
            if len(comp) <= orig_comp:
                new_raw = pcode + comp
            else:
                print(f"    Still overflow ({len(comp)} > {orig_comp}). Skipping.")
                continue
        ok = patch_ole_stream(vba_bin, ole, path, new_raw)
        print(f"    Patched: {ok}")

    # ── Rebuild dir stream with UserForm2 -> FormNodeSelect ──
    print("\nRebuilding dir stream...")
    dir_orig = get_stream(ole, 'VBA/dir')
    dir_new = rename_in_dir_stream(dir_orig, 'UserForm2', 'FormNodeSelect')
    print(f"  dir: {len(dir_orig)} -> {len(dir_new)} bytes")
    ok = patch_ole_stream(vba_bin, ole, 'VBA/dir', dir_new)
    print(f"  dir patched: {ok}")

    # ── Update PROJECT stream ──
    print("\nUpdating PROJECT stream...")
    proj_orig = get_stream(ole, 'PROJECT')
    proj_text = proj_orig.decode('latin-1')
    proj_new = proj_text.replace('BaseClass=UserForm2', 'BaseClass=FormNodeSelect')
    # Remove the UserForm2 workspace-position line (optional IDE metadata)
    proj_new = re.sub(r'\r?\nUserForm2=[^\n]*', '', proj_new)
    proj_new_bytes = proj_new.encode('latin-1')
    print(f"  PROJECT: {len(proj_orig)} -> {len(proj_new_bytes)} bytes")
    ok = patch_ole_stream(vba_bin, ole, 'PROJECT', proj_new_bytes)
    print(f"  PROJECT patched: {ok}")

    # ── Patch OLE directory entries for UserForm2 -> FormNodeSelect ──
    print("\nPatching OLE directory entries...")
    patch_ole_dir_entries(vba_bin, 'UserForm2', 'FormNodeSelect')

    # ── Write patched xlsm ──
    out_buf = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(xlsm_orig), 'r') as zin2:
        with zipfile.ZipFile(out_buf, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin2.infolist():
                if item.filename == 'xl/vbaProject.bin':
                    zout.writestr(item, bytes(vba_bin))
                else:
                    zout.writestr(item, zin2.read(item.filename))

    with open(XLSM_PATH, 'wb') as f:
        f.write(out_buf.getvalue())
    print(f"\nDone. Wrote {XLSM_PATH}")


if __name__ == '__main__':
    main()
