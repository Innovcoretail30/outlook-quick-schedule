#!/usr/bin/env node
import fs from "node:fs";
import path from "node:path";
import zlib from "node:zlib";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const outDir = path.join(root, "src/assets");
fs.mkdirSync(outDir, { recursive: true });

function crc32(buf) {
  let c = ~0;
  for (let i = 0; i < buf.length; i++) {
    c ^= buf[i];
    for (let k = 0; k < 8; k++) {
      c = (c >>> 1) ^ (0xedb88320 & -(c & 1));
    }
  }
  return ~c >>> 0;
}

function chunk(type, data) {
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length);
  const td = Buffer.concat([Buffer.from(type), data]);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(td));
  return Buffer.concat([len, td, crc]);
}

function encodePng(width, height, pixels) {
  const raw = Buffer.alloc((width * 4 + 1) * height);
  for (let y = 0; y < height; y++) {
    raw[(width * 4 + 1) * y] = 0;
    for (let x = 0; x < width; x++) {
      const i = (y * width + x) * 4;
      const o = (width * 4 + 1) * y + 1 + x * 4;
      raw[o] = pixels[i];
      raw[o + 1] = pixels[i + 1];
      raw[o + 2] = pixels[i + 2];
      raw[o + 3] = pixels[i + 3];
    }
  }
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8;
  ihdr[9] = 6;
  return Buffer.concat([
    Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]),
    chunk("IHDR", ihdr),
    chunk("IDAT", zlib.deflateSync(raw)),
    chunk("IEND", Buffer.alloc(0))
  ]);
}

function canvas(size) {
  const pixels = Buffer.alloc(size * size * 4);
  return {
    size,
    pixels,
    set(x, y, r, g, b, a) {
      if (x < 0 || y < 0 || x >= size || y >= size) {
        return;
      }
      const i = (y * size + x) * 4;
      const na = a / 255;
      const oa = pixels[i + 3] / 255;
      const outA = na + oa * (1 - na);
      if (outA <= 0) {
        return;
      }
      pixels[i] = Math.round((r * na + pixels[i] * oa * (1 - na)) / outA);
      pixels[i + 1] = Math.round((g * na + pixels[i + 1] * oa * (1 - na)) / outA);
      pixels[i + 2] = Math.round((b * na + pixels[i + 2] * oa * (1 - na)) / outA);
      pixels[i + 3] = Math.round(outA * 255);
    }
  };
}

function fillCircle(c, cx, cy, radius, r, g, b, a) {
  const size = c.size;
  const r2 = radius * radius;
  const rOuter = radius + 1.2;
  const rOuter2 = rOuter * rOuter;
  const x0 = Math.max(0, Math.floor(cx - rOuter));
  const x1 = Math.min(size - 1, Math.ceil(cx + rOuter));
  const y0 = Math.max(0, Math.floor(cy - rOuter));
  const y1 = Math.min(size - 1, Math.ceil(cy + rOuter));
  for (let y = y0; y <= y1; y++) {
    for (let x = x0; x <= x1; x++) {
      const dx = x + 0.5 - cx;
      const dy = y + 0.5 - cy;
      const d2 = dx * dx + dy * dy;
      if (d2 <= r2) {
        c.set(x, y, r, g, b, a);
      } else if (d2 <= rOuter2) {
        const d = Math.sqrt(d2);
        const t = Math.max(0, 1 - (d - radius));
        c.set(x, y, r, g, b, Math.round(a * t));
      }
    }
  }
}

function strokeCircle(c, cx, cy, radius, width, r, g, b, a) {
  const size = c.size;
  const inner = radius - width / 2;
  const outer = radius + width / 2;
  const x0 = Math.max(0, Math.floor(cx - outer - 1));
  const x1 = Math.min(size - 1, Math.ceil(cx + outer + 1));
  const y0 = Math.max(0, Math.floor(cy - outer - 1));
  const y1 = Math.min(size - 1, Math.ceil(cy + outer + 1));
  for (let y = y0; y <= y1; y++) {
    for (let x = x0; x <= x1; x++) {
      const dx = x + 0.5 - cx;
      const dy = y + 0.5 - cy;
      const d = Math.sqrt(dx * dx + dy * dy);
      const dist = Math.abs(d - radius);
      if (dist <= width / 2) {
        c.set(x, y, r, g, b, a);
      } else if (dist <= width / 2 + 1) {
        const t = 1 - (dist - width / 2);
        c.set(x, y, r, g, b, Math.round(a * t));
      }
    }
  }
}

function fillRect(c, x0, y0, x1, y1, r, g, b, a) {
  for (let y = Math.floor(y0); y <= Math.ceil(y1); y++) {
    for (let x = Math.floor(x0); x <= Math.ceil(x1); x++) {
      c.set(x, y, r, g, b, a);
    }
  }
}

function drawHand(c, cx, cy, angle, length, width, r, g, b) {
  const steps = Math.ceil(length + 4);
  for (let i = 0; i <= steps; i++) {
    const t = i / steps;
    const x = cx + Math.sin(angle) * length * t;
    const y = cy - Math.cos(angle) * length * t;
    fillCircle(c, x, y, width / 2, r, g, b, 255);
  }
}

function drawClock(c, color, face) {
  const s = c.size;
  const cx = s / 2;
  const cy = s / 2;
  const radius = s * 0.42;
  fillCircle(c, cx, cy, radius, color[0], color[1], color[2], 255);
  if (face) {
    fillCircle(c, cx, cy, radius * 0.78, face[0], face[1], face[2], 255);
  }
  const hand = face ? color : [255, 255, 255];
  drawHand(c, cx, cy, Math.PI * 0.15, radius * 0.42, s * 0.07, hand[0], hand[1], hand[2]);
  drawHand(c, cx, cy, Math.PI * 0.85, radius * 0.28, s * 0.09, hand[0], hand[1], hand[2]);
  fillCircle(c, cx, cy, s * 0.05, hand[0], hand[1], hand[2], 255);
}

function drawGear(c, color) {
  const s = c.size;
  const cx = s / 2;
  const cy = s / 2;
  const teeth = 8;
  const outer = s * 0.44;
  const inner = s * 0.30;
  const hole = s * 0.12;
  for (let y = 0; y < s; y++) {
    for (let x = 0; x < s; x++) {
      const dx = x + 0.5 - cx;
      const dy = y + 0.5 - cy;
      const d = Math.sqrt(dx * dx + dy * dy);
      const angle = Math.atan2(dy, dx);
      const tooth = (Math.cos(angle * teeth) + 1) / 2;
      const rim = inner + (outer - inner) * (tooth > 0.55 ? 1 : 0.15);
      if (d <= rim && d >= hole) {
        c.set(x, y, color[0], color[1], color[2], 255);
      }
    }
  }
}

const BLUE = [15, 108, 189];
const WHITE = [255, 255, 255];

function save(name, c) {
  fs.writeFileSync(path.join(outDir, name), encodePng(c.size, c.size, c.pixels));
}

for (const size of [16, 32, 64, 80, 128]) {
  const clock = canvas(size);
  drawClock(clock, BLUE, WHITE);
  save("icon-" + size + ".png", clock);

  const gear = canvas(size);
  drawGear(gear, BLUE);
  save("icon-gear-" + size + ".png", gear);
}

const color = canvas(192);
drawClock(color, BLUE, WHITE);
save("color.png", color);

const outline = canvas(32);
drawClock(outline, WHITE, null);
save("outline.png", outline);

const fav = canvas(32);
drawClock(fav, BLUE, WHITE);
save("favicon.png", fav);

console.log("Wrote icons to src/assets");
