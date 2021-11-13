""" inspired by: http://b2xtranslator.sourceforge.net/howtos/How_to_retrieve_text_from_a_binary_doc_file.pdf """
import struct
from typing import List

import olefile
from olefile.olefile import OleFileIO, OleStream


class OLETextExtract:
    def __init__(self):
        pass

    def get_uint32(self, data: bytes, offset: int) -> int:
        return struct.unpack("<I", data[offset : offset + 4])[0]

    def table_stream_name(self, fib: bytes) -> str:
        bit = (fib[0xB] >> 1) & 1
        return f"{bit}Table"

    def _load_table(self, ole: OleFileIO, fib: bytes) -> bytes:
        offset: int = self.get_uint32(fib, 0x1A2)
        length: int = self.get_uint32(fib, 0x1A6)
        table_name: str = self.table_stream_name(fib)

        s: OleStream = ole.openstream(table_name)
        s.seek(offset)
        table: bytes = s.read(length)
        return table

    def _load_piece_table(self, table: bytes) -> bytes:
        i = 0
        while i < len(table):
            entry_type: int = table[i]
            if entry_type == 1:
                i += 2 + table[i + 1]
            elif entry_type == 2:
                piece_table_length: int = self.get_uint32(table, i + 1)
                piece_table: bytes = table[i + 5 : i + 5 + piece_table_length]
                return piece_table
            else:
                return

    def _get_text(self, doc: bytes, piece_table: bytes) -> str:
        piece_count: int = (len(piece_table) - 4) // 12
        character_positions: List[int] = []
        for i in range(piece_count + 1):
            character_positions.append(self.get_uint32(piece_table, i * 4))

        text: str = ""
        for i in range(piece_count):
            cp_start: int = character_positions[i]
            cp_end: int = character_positions[i + 1]
            desc_offset: int = (piece_count + 1) * 4 + i * 8
            descriptor: bytes = piece_table[desc_offset : desc_offset + 8]
            fc_value: int = self.get_uint32(descriptor, 2)
            is_ansi: bool = (fc_value & 0x40000000) == 0x40000000
            fc: int = fc_value & 0xBFFFFFFF
            cb: int = cp_end - cp_start
            if is_ansi:
                fc = fc // 2
                encoding: str = "cp1252"
            else:
                fc = fc
                encoding: str = "utf16"
                cb *= 2
            raw: bytes = doc[fc : fc + cb]
            if is_ansi:
                raw = raw.replace(b"\r", b"\n")
            else:
                raw = raw.replace(b"\x00\r", b"\x00\n")
            text += raw.decode(encoding)
        return text

    def extract(self, path: str) -> str:
        with olefile.OleFileIO(path) as ole:
            s: OleFileIO = ole.openstream("WordDocument")
            doc: bytes = s.read()
            fib: bytes = doc[:1472]
            table: bytes = self._load_table(ole, fib)
            piece_table: bytes = self._load_piece_table(table)
            text: str = self._get_text(doc, piece_table)
            return text
