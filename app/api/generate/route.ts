import { NextResponse } from 'next/server';
import { GoogleGenerativeAI } from '@google/generative-ai';
import pptxgen from 'pptxgenjs';

const genAI = new GoogleGenerativeAI(process.env.GEMINI_API_KEY!);
const PEXELS_KEY = process.env.PEXELS_API_KEY!;

// Fungsi ambil gambar dari Pexels
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

// Fungsi download gambar jadi buffer
async function downloadImage(url: string): Promise<Buffer | null> {
  try {
    const res = await fetch(url);
    const arrayBuffer = await res.arrayBuffer();
    return Buffer.from(arrayBuffer);
  } catch {
    return null;
  }
}

export async function POST(req: Request) {
  const { prompt } = await req.json();

  if (!prompt) {
    return NextResponse.json({ error: 'Prompt kosong' }, { status: 400 });
  }

  try {
    const model = genAI.getGenerativeModel({ model: 'gemini-3.6-flash' });

    const systemPrompt = `
      Buat struktur presentasi PowerPoint tentang "${prompt}".
      Output HARUS JSON valid, tanpa markdown, tanpa code block.
      Format:
      {
        "slides": [
          {
            "layout": "title" | "content" | "image-left" | "image-right" | "full-image",
            "title": "Judul Slide",
            "bullets": ["poin 1", "poin 2", "poin 3"],
            "image_keyword": "english keyword untuk cari foto stock"
          }
        ]
      }
      
      Aturan:
      - Slide pertama WAJIB layout "title" (judul presentasi + subtitle)
      - Slide terakhir layout "content" (kesimpulan)
      - Slide tengah variasi: "image-left", "image-right", atau "content"
      - image_keyword HARUS bahasa Inggris, 2-3 kata, relevan dengan topik
      - Setiap slide maksimal 4 bullet
      - Buat 5-7 slide total
    `;

    const result = await model.generateContent(systemPrompt);
    const text = result.response.text();
    const cleaned = text.replace(/```json/g, '').replace(/```/g, '').trim();
    const data = JSON.parse(cleaned);

    // Generate PPTX
    const pptx = new pptxgen();
    pptx.layout = 'LAYOUT_16x9';
    pptx.author = 'PPT AI Generator';
    pptx.title = prompt;

    for (const slideData of data.slides) {
      const slide = pptx.addSlide();
      const layout = slideData.layout || 'content';

      // Ambil gambar kalau ada keyword
      let imageBuffer: Buffer | null = null;
      if (slideData.image_keyword) {
        const imageUrl = await getImage(slideData.image_keyword);
        if (imageUrl) {
          imageBuffer = await downloadImage(imageUrl);
        }
      }

      if (layout === 'title') {
        // Layout judul: background gelap + teks putih
        slide.background = { color: '1E293B' };
        slide.addText(slideData.title, {
          x: 0.5, y: 2.2, w: '90%', h: 1.5,
          fontSize: 40, bold: true, color: 'FFFFFF',
          align: 'center', valign: 'middle'
        });
        if (slideData.bullets?.[0]) {
          slide.addText(slideData.bullets[0], {
            x: 0.5, y: 3.8, w: '90%', h: 0.8,
            fontSize: 18, color: 'CBD5E1', align: 'center'
          });
        }
      } else if (layout === 'image-left' && imageBuffer) {
        // Gambar di kiri, teks di kanan
        slide.addImage({
          data: `image/png;base64,${imageBuffer.toString('base64')}`,
          x: 0, y: 0, w: 5, h: 5.63,
        });
        slide.addText(slideData.title, {
          x: 5.3, y: 0.5, w: 4.5, h: 0.9,
          fontSize: 26, bold: true, color: '1E293B'
        });
        slide.addText(
          slideData.bullets.map((b: string) => ({
            text: b,
            options: { bullet: true, fontSize: 16, color: '475569', breakLine: true }
          })),
          { x: 5.3, y: 1.6, w: 4.5, h: 3.5, valign: 'top' }
        );
      } else if (layout === 'image-right' && imageBuffer) {
        // Teks di kiri, gambar di kanan
        slide.addText(slideData.title, {
          x: 0.5, y: 0.5, w: 4.5, h: 0.9,
          fontSize: 26, bold: true, color: '1E293B'
        });
        slide.addText(
          slideData.bullets.map((b: string) => ({
            text: b,
            options: { bullet: true, fontSize: 16, color: '475569', breakLine: true }
          })),
          { x: 0.5, y: 1.6, w: 4.5, h: 3.5, valign: 'top' }
        );
        slide.addImage({
          data: `image/png;base64,${imageBuffer.toString('base64')}`,
          x: 5, y: 0, w: 5, h: 5.63,
        });
      } else if (layout === 'full-image' && imageBuffer) {
        // Background full image + overlay
        slide.addImage({
          data: `image/png;base64,${imageBuffer.toString('base64')}`,
          x: 0, y: 0, w: 10, h: 5.63,
        });
        slide.addShape('rect', {
          x: 0, y: 3.5, w: 10, h: 2.13,
          fill: { color: '000000', transparency: 30 }
        });
        slide.addText(slideData.title, {
          x: 0.5, y: 3.8, w: 9, h: 1.5,
          fontSize: 28, bold: true, color: 'FFFFFF'
        });
      } else {
        // Layout content default
        slide.addText(slideData.title, {
          x: 0.5, y: 0.5, w: '90%', h: 1,
          fontSize: 28, bold: true, color: '1E293B'
        });
        slide.addText(
          slideData.bullets.map((b: string) => ({
            text: b,
            options: { bullet: true, fontSize: 18, color: '475569', breakLine: true }
          })),
          { x: 0.5, y: 1.8, w: '90%', h: '70%', valign: 'top' }
        );
      }
    }

    const buffer = (await pptx.write({ outputType: 'nodebuffer' })) as Buffer;

    return new NextResponse(new Uint8Array(buffer), {
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        'Content-Disposition': `attachment; filename="presentasi.pptx"`,
      },
    });
  } catch (error) {
    console.error('Error generating PPT:', error);
    return NextResponse.json({ error: 'Gagal generate PPT' }, { status: 500 });
  }
}

export const runtime = 'nodejs';
export const maxDuration = 60;
