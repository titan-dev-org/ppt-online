import { NextResponse } from 'next/server';
import { GoogleGenerativeAI } from '@google/generative-ai';
import pptxgen from 'pptxgenjs';

const genAI = new GoogleGenerativeAI(process.env.GEMINI_API_KEY!);
const PEXELS_KEY = process.env.PEXELS_API_KEY!;

// 🎨 Palet warna berdasarkan topik
const PALETTES: Record<string, { primary: string; secondary: string; accent: string; bg: string; text: string }> = {
  health:   { primary: '059669', secondary: '10B981', accent: '34D399', bg: 'ECFDF5', text: '064E3B' },
  tech:     { primary: '2563EB', secondary: '3B82F6', accent: '60A5FA', bg: 'EFF6FF', text: '1E3A8A' },
  business: { primary: 'D97706', secondary: 'F59E0B', accent: 'FBBF24', bg: 'FFFBEB', text: '78350F' },
  education:{ primary: '7C3AED', secondary: '8B5CF6', accent: 'A78BFA', bg: 'F5F3FF', text: '4C1D95' },
  nature:   { primary: '16A34A', secondary: '22C55E', accent: '4ADE80', bg: 'F0FDF4', text: '14532D' },
  default:  { primary: '1E293B', secondary: '334155', accent: '64748B', bg: 'F8FAFC', text: '0F172A' },
};

function pickPalette(topic: string) {
  const t = topic.toLowerCase();
  if (/sehat|kesehatan|medis|dokter|nutrisi|metabol/.test(t)) return PALETTES.health;
  if (/teknologi|ai|komputer|digital|software|data/.test(t)) return PALETTES.tech;
  if (/bisnis|keuangan|marketing|startup|investasi/.test(t)) return PALETTES.business;
  if (/sekolah|belajar|pendidikan|siswa|kuliah/.test(t)) return PALETTES.education;
  if (/alam|lingkungan|bumi|hijau|climate/.test(t)) return PALETTES.nature;
  return PALETTES.default;
}

async function getImage(keyword: string): Promise<string | null> {
  try {
    const res = await fetch(
      `https://api.pexels.com/v1/search?query=${encodeURIComponent(keyword)}&per_page=1&orientation=landscape`,
      { headers: { Authorization: PEXELS_KEY } }
    );
    const data = await res.json();
    return data.photos?.[0]?.src?.large2x || null;
  } catch {
    return null;
  }
}

async function downloadImage(url: string): Promise<Buffer | null> {
  try {
    const res = await fetch(url);
    return Buffer.from(await res.arrayBuffer());
  } catch {
    return null;
  }
}

export async function POST(req: Request) {
  const { prompt } = await req.json();
  if (!prompt) return NextResponse.json({ error: 'Prompt kosong' }, { status: 400 });

  try {
    const model = genAI.getGenerativeModel({ model: 'gemini-3.6-flash' });

    const systemPrompt = `
      Buat struktur presentasi PowerPoint tentang "${prompt}".
      Output HARUS JSON valid tanpa markdown.
      Format:
      {
        "slides": [
          {
            "layout": "title" | "content" | "image-left" | "image-right" | "full-image" | "quote" | "stats",
            "title": "Judul",
            "subtitle": "subjudul (opsional, buat title/quote)",
            "bullets": ["poin 1", "poin 2"],
            "image_keyword": "english keyword 2-3 kata",
            "stat": { "value": "75%", "label": "keterangan" }
          }
        ]
      }
      Aturan:
      - Slide 1: layout "title"
      - Slide terakhir: layout "content" (kesimpulan)
      - Tengah: variasi layout
      - image_keyword HARUS bahasa Inggris
      - Maks 4 bullet per slide
      - 6-8 slide total
    `;

    const result = await model.generateContent(systemPrompt);
    const text = result.response.text();
    const cleaned = text.replace(/```json/g, '').replace(/```/g, '').trim();
    const data = JSON.parse(cleaned);

    const palette = pickPalette(prompt);
    const totalSlides = data.slides.length;

    const pptx = new pptxgen();
    pptx.layout = 'LAYOUT_16x9';
    pptx.author = 'PPT AI Generator';
    pptx.title = prompt;

    // Fetch semua gambar parallel
    const images = await Promise.all(
      data.slides.map(async (s: any) => {
        if (!s.image_keyword) return null;
        const url = await getImage(s.image_keyword);
        return url ? await downloadImage(url) : null;
      })
    );

    // 🎨 Fungsi dekorasi background
    const addDecorations = (slide: any, isDark = false) => {
      // Shape dekoratif di pojok kanan atas
      slide.addShape('ellipse', {
        x: 8.5, y: -1, w: 3, h: 3,
        fill: { color: isDark ? palette.secondary : palette.accent, transparency: 70 },
        line: { color: palette.accent, width: 0 }
      });
      // Shape dekoratif di pojok kiri bawah
      slide.addShape('ellipse', {
        x: -1, y: 4.5, w: 2.5, h: 2.5,
        fill: { color: isDark ? palette.accent : palette.secondary, transparency: 80 },
        line: { color: palette.accent, width: 0 }
      });
    };

    // 🎨 Accent bar di bawah judul
    const addAccentBar = (slide: any, y = 1.3) => {
      slide.addShape('rect', {
        x: 0.5, y, w: 1.2, h: 0.08,
        fill: { color: palette.primary }
      });
    };

    // 🎨 Footer + page number
    const addFooter = (slide: any, idx: number) => {
      slide.addText(prompt.slice(0, 40), {
        x: 0.5, y: 5.25, w: 6, h: 0.3,
        fontSize: 9, color: palette.text, transparency: 50
      });
      slide.addText(`${idx + 1} / ${totalSlides}`, {
        x: 8.5, y: 5.25, w: 1, h: 0.3,
        fontSize: 9, color: palette.text, align: 'right', transparency: 50
      });
    };

    for (let i = 0; i < data.slides.length; i++) {
      const s = data.slides[i];
      const layout = s.layout || 'content';
      const img = images[i];
      const slide = pptx.addSlide();

      // Background dasar semua slide
      slide.background = { color: palette.bg };

      if (layout === 'title') {
        // 🌟 Slide judul dengan background gelap + dekorasi
        slide.background = { color: palette.text };
        addDecorations(slide, true);

        slide.addText(s.title, {
          x: 0.7, y: 2.0, w: '85%', h: 1.5,
          fontSize: 44, bold: true, color: 'FFFFFF',
          align: 'center', valign: 'middle', fontFace: 'Arial'
        });

        if (s.subtitle || s.bullets?.[0]) {
          slide.addText(s.subtitle || s.bullets[0], {
            x: 0.7, y: 3.6, w: '85%', h: 0.8,
            fontSize: 18, color: palette.accent,
            align: 'center', italic: true
          });
        }

        // Garis aksen di bawah judul
        slide.addShape('rect', {
          x: 4.5, y: 3.5, w: 1, h: 0.05,
          fill: { color: palette.accent }
        });
      }
      else if (layout === 'image-left' && img) {
        slide.addImage({
          data: `image/png;base64,${img.toString('base64')}`,
          x: 0, y: 0, w: 4.8, h: 5.63,
        });
        // Overlay gradient tipis di atas gambar
        slide.addShape('rect', {
          x: 0, y: 0, w: 4.8, h: 5.63,
          fill: { color: palette.primary, transparency: 75 }
        });

        addDecorations(slide);
        slide.addText(s.title, {
          x: 5.2, y: 0.6, w: 4.4, h: 1,
          fontSize: 26, bold: true, color: palette.text
        });
        addAccentBar(slide, 1.7);
        slide.addText(
          s.bullets.map((b: string) => ({
            text: b,
            options: { bullet: { code: '25CF' }, fontSize: 15, color: palette.text, breakLine: true, paraSpaceAfter: 8 }
          })),
          { x: 5.2, y: 1.9, w: 4.4, h: 3.2, valign: 'top' }
        );
        addFooter(slide, i);
      }
      else if (layout === 'image-right' && img) {
        addDecorations(slide);
        slide.addText(s.title, {
