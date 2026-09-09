/*!
 * zip.js — 零依赖 ZIP 读取器（支持 stored / deflate）
 * ------------------------------------------------------------------
 * 用于在浏览器与 Node 中读取 .docx（本质是 ZIP 包）内的 XML 部件。
 * 不依赖 JSZip，完全离线可用。
 *
 * 对外 API：
 *   WordMd.zip.unzip(ArrayBuffer|Uint8Array) -> { "word/document.xml": Uint8Array, ... }
 *   WordMd.zip.text(entry) -> string (UTF-8 解码)
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory();
  } else {
    root.WordMd = root.WordMd || {};
    root.WordMd.zip = factory();
  }
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';

  /* ============================ 位读取器 ============================ */

  function BitReader(data, pos) {
    this.data = data;
    this.pos = pos;
    this.buf = 0;
    this.cnt = 0;
  }
  BitReader.prototype.readBit = function () {
    if (this.cnt === 0) {
      this.buf = this.data[this.pos++];
      this.cnt = 8;
    }
    var b = this.buf & 1;
    this.buf >>= 1;
    this.cnt--;
    return b;
  };
  BitReader.prototype.readBits = function (n) {
    var v = 0;
    for (var i = 0; i < n; i++) v |= this.readBit() << i;
    return v;
  };
  BitReader.prototype.alignByte = function () {
    this.cnt = 0;
    this.buf = 0;
  };

  /* ============================ 动态字节写入 ============================ */

  function ByteWriter(initial) {
    this.buf = new Uint8Array(initial || 65536);
    this.len = 0;
  }
  ByteWriter.prototype.ensure = function (extra) {
    if (this.len + extra <= this.buf.length) return;
    var size = this.buf.length;
    while (size < this.len + extra) size *= 2;
    var next = new Uint8Array(size);
    next.set(this.buf.subarray(0, this.len));
    this.buf = next;
  };
  ByteWriter.prototype.push = function (byte) {
    this.ensure(1);
    this.buf[this.len++] = byte & 0xff;
  };
  ByteWriter.prototype.copyFrom = function (src, offset, length) {
    this.ensure(length);
    for (var i = 0; i < length; i++) this.buf[this.len++] = src[offset + i];
  };
  ByteWriter.prototype.copyWithinBack = function (distance, length) {
    this.ensure(length);
    var start = this.len - distance;
    for (var i = 0; i < length; i++) this.buf[this.len++] = this.buf[start + i];
  };
  ByteWriter.prototype.toUint8Array = function () {
    return this.buf.subarray(0, this.len);
  };

  /* ============================ Huffman 表 ============================ */

  function buildHuffman(lengths, count) {
    var counts = new Int32Array(16);
    var i;
    for (i = 0; i < count; i++) counts[lengths[i]]++;
    counts[0] = 0;

    var offsets = new Int32Array(16);
    for (i = 1; i < 16; i++) offsets[i] = offsets[i - 1] + counts[i - 1];

    var symbols = new Int32Array(count);
    for (i = 0; i < count; i++) {
      if (lengths[i]) symbols[offsets[lengths[i]]++] = i;
    }
    return { counts: counts, symbols: symbols };
  }

  function decodeSymbol(br, huff) {
    var code = 0, first = 0, index = 0;
    for (var len = 1; len <= 15; len++) {
      code |= br.readBit();
      var count = huff.counts[len];
      if (code - first < count) return huff.symbols[index + (code - first)];
      index += count;
      first = (first + count) << 1;
      code <<= 1;
    }
    throw new Error('ZIP: 无效的 Huffman 编码');
  }

  var LENGTH_BASE = [3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258];
  var LENGTH_EXTRA = [0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0];
  var DIST_BASE = [1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577];
  var DIST_EXTRA = [0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13];
  var CL_ORDER = [16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15];

  var FIXED_LIT = (function () {
    var lengths = new Uint8Array(288);
    var i;
    for (i = 0; i < 144; i++) lengths[i] = 8;
    for (; i < 256; i++) lengths[i] = 9;
    for (; i < 280; i++) lengths[i] = 7;
    for (; i < 288; i++) lengths[i] = 8;
    return buildHuffman(lengths, 288);
  })();

  var FIXED_DIST = (function () {
    var lengths = new Uint8Array(30);
    for (var i = 0; i < 30; i++) lengths[i] = 5;
    return buildHuffman(lengths, 30);
  })();

  /* ============================ inflate（raw deflate）============================ */

  /**
   * 单个压缩条目的解压输出上限。
   * 防御「解压炸弹」：几 KB 的压缩数据可以膨胀到数 GB，耗尽浏览器内存。
   */
  var MAX_INFLATE_BYTES = 200 * 1024 * 1024;

  function inflateRaw(data, offset) {
    var br = new BitReader(data, offset || 0);
    var out = new ByteWriter(Math.max(4096, data.length * 4));
    var final = 0;

    function guard() {
      if (out.len > MAX_INFLATE_BYTES) {
        throw new Error('ZIP: 解压数据超过安全上限（' + (MAX_INFLATE_BYTES / 1024 / 1024) + ' MB），已中止');
      }
    }

    do {
      final = br.readBit();
      var type = br.readBits(2);

      if (type === 0) {
        // 未压缩块
        br.alignByte();
        var p = br.pos;
        var len = data[p] | (data[p + 1] << 8);
        br.pos = p + 4;
        out.copyFrom(data, br.pos, len);
        br.pos += len;
      } else if (type === 1) {
        inflateBlock(br, out, FIXED_LIT, FIXED_DIST);
      } else if (type === 2) {
        var hlit = br.readBits(5) + 257;
        var hdist = br.readBits(5) + 1;
        var hclen = br.readBits(4) + 4;

        var clLengths = new Uint8Array(19);
        for (var i = 0; i < hclen; i++) clLengths[CL_ORDER[i]] = br.readBits(3);
        var clHuff = buildHuffman(clLengths, 19);

        var lengths = new Uint8Array(hlit + hdist);
        var n = 0;
        while (n < hlit + hdist) {
          var sym = decodeSymbol(br, clHuff);
          if (sym < 16) {
            lengths[n++] = sym;
          } else if (sym === 16) {
            var prev = lengths[n - 1];
            var rep = 3 + br.readBits(2);
            while (rep-- > 0) lengths[n++] = prev;
          } else if (sym === 17) {
            var r17 = 3 + br.readBits(3);
            while (r17-- > 0) lengths[n++] = 0;
          } else {
            var r18 = 11 + br.readBits(7);
            while (r18-- > 0) lengths[n++] = 0;
          }
        }
        var litHuff = buildHuffman(lengths.subarray(0, hlit), hlit);
        var distHuff = buildHuffman(lengths.subarray(hlit), hdist);
        inflateBlock(br, out, litHuff, distHuff);
      } else {
        throw new Error('ZIP: 非法的压缩块类型');
      }
      guard();
    } while (!final);

    return out.toUint8Array();
  }

  function inflateBlock(br, out, litHuff, distHuff) {
    for (;;) {
      var sym = decodeSymbol(br, litHuff);
      if (sym < 256) {
        out.push(sym);
      } else if (sym === 256) {
        return;
      } else {
        var li = sym - 257;
        if (li >= LENGTH_BASE.length) throw new Error('ZIP: 非法长度码');
        var length = LENGTH_BASE[li] + br.readBits(LENGTH_EXTRA[li]);
        var dsym = decodeSymbol(br, distHuff);
        if (dsym >= DIST_BASE.length) throw new Error('ZIP: 非法距离码');
        var distance = DIST_BASE[dsym] + br.readBits(DIST_EXTRA[dsym]);
        if (distance > out.len) throw new Error('ZIP: 距离超出输出缓冲');
        out.copyWithinBack(distance, length);
        if (out.len > MAX_INFLATE_BYTES) {
          throw new Error('ZIP: 解压数据超过安全上限（' + (MAX_INFLATE_BYTES / 1024 / 1024) + ' MB），已中止');
        }
      }
    }
  }

  /* ============================ ZIP 容器解析 ============================ */

  function findEOCD(view, length) {
    var min = Math.max(0, length - 65557);
    for (var i = length - 22; i >= min; i--) {
      if (view[i] === 0x50 && view[i + 1] === 0x4b && view[i + 2] === 0x05 && view[i + 3] === 0x06) {
        return i;
      }
    }
    return -1;
  }

  function utf8Decode(bytes) {
    if (typeof TextDecoder !== 'undefined') {
      return new TextDecoder('utf-8').decode(bytes);
    }
    var s = '';
    for (var i = 0; i < bytes.length; i++) {
      var c = bytes[i];
      if (c < 0x80) s += String.fromCharCode(c);
      else if (c < 0xe0) s += String.fromCharCode(((c & 0x1f) << 6) | (bytes[++i] & 0x3f));
      else if (c < 0xf0) {
        s += String.fromCharCode(((c & 0x0f) << 12) | ((bytes[++i] & 0x3f) << 6) | (bytes[++i] & 0x3f));
      } else {
        var cp = ((c & 0x07) << 18) | ((bytes[++i] & 0x3f) << 12) | ((bytes[++i] & 0x3f) << 6) | (bytes[++i] & 0x3f);
        cp -= 0x10000;
        s += String.fromCharCode(0xd800 + (cp >> 10), 0xdc00 + (cp & 0x3ff));
      }
    }
    return s;
  }

  function unzip(input) {
    var bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
    var view = bytes;
    var eocd = findEOCD(view, bytes.length);
    if (eocd < 0) throw new Error('ZIP: 找不到中央目录，文件可能损坏或不是 docx');

    function u16(o) { return view[o] | (view[o + 1] << 8); }
    function u32(o) {
      return (view[o] | (view[o + 1] << 8) | (view[o + 2] << 16) | (view[o + 3] << 24)) >>> 0;
    }

    var entryCount = u16(eocd + 10);
    var cdOffset = u32(eocd + 16);
    var files = {};
    var p = cdOffset;

    for (var i = 0; i < entryCount; i++) {
      if (u32(p) !== 0x02014b50) break;
      var method = u16(p + 10);
      var compSize = u32(p + 20);
      var rawSize = u32(p + 24);
      var nameLen = u16(p + 28);
      var extraLen = u16(p + 30);
      var commentLen = u16(p + 32);
      var localOffset = u32(p + 42);
      var nameBytes = bytes.subarray(p + 46, p + 46 + nameLen);
      var name = utf8Decode(nameBytes);

      // 读取本地头以定位数据起点
      if (u32(localOffset) !== 0x04034b50) throw new Error('ZIP: 本地文件头损坏');
      var lNameLen = u16(localOffset + 26);
      var lExtraLen = u16(localOffset + 28);
      var dataStart = localOffset + 30 + lNameLen + lExtraLen;

      if (!/\/$/.test(name)) {
        var raw;
        if (method === 0) {
          raw = bytes.slice(dataStart, dataStart + rawSize);
        } else if (method === 8) {
          raw = inflateRaw(bytes, dataStart);
        } else {
          throw new Error('ZIP: 不支持的压缩方式 ' + method + '（' + name + '）');
        }
        files[name] = raw;
      }

      p += 46 + nameLen + extraLen + commentLen;
    }
    return files;
  }

  function text(entry) {
    if (!entry) return '';
    var bytes = entry instanceof Uint8Array ? entry : new Uint8Array(entry);
    // 去掉 UTF-8 BOM
    if (bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) {
      bytes = bytes.subarray(3);
    }
    return utf8Decode(bytes);
  }

  /* ============================ CRC32（校验用）============================ */

  var CRC_TABLE = (function () {
    var t = new Int32Array(256);
    for (var n = 0; n < 256; n++) {
      var c = n;
      for (var k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      t[n] = c;
    }
    return t;
  })();

  function crc32(bytes) {
    var c = 0 ^ -1;
    for (var i = 0; i < bytes.length; i++) c = (c >>> 8) ^ CRC_TABLE[(c ^ bytes[i]) & 0xff];
    return (c ^ -1) >>> 0;
  }

  return {
    unzip: unzip,
    text: text,
    crc32: crc32,
    inflateRaw: inflateRaw
  };
});
