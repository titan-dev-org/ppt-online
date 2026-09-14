import { NextResponse } from 'next/server';
import { GoogleGenerativeAI } from '@google/generative-ai';
import pptxgen from 'pptxgenjs';

// Inisialisasi Gemini API
const genAI = new GoogleGenerativeAI(process.env.GEMINI_API_KEY!);

export async function POST(req: Request) {
  const { prompt } = await req.json();

  if (!prompt) {
    return NextResponse.json({ error: 'Prompt kosong' }, { status: 400 });
  }

  try {
    // Gunakan model gemini-3.8-flash
    const model = genAI.getGenerativeModel({ model: 'gemini-3.8-flash' });

    // System prompt untuk generate struktur PPT
    const systemPrompt = `
      Buat struktur presentasi PowerPoint tentang "${prompt}".
      Output HARUS JSON valid, tanpa markdown, tanpa code block.
      Format:
      {
        "slides": [
          {
            "title": "Judul Slide",
            "bullets": ["poin 1", "poin 2", "poin 3"]
          }
        ]
      }
      Buat 4-6 slide. Setiap slide maksimal 4 bullet.
    `;

    const result = await model.generateContent(systemPrompt);
    const text = result.response.text();

    // Bersihin kemungkinan AI ngasih markdown
    const cleaned = text.replace(/```json/g, '').replace(/```/g, '').trim();
    const data = JSON.parse(cleaned);

    // Generate file PPTX
    const pptx = new pptxgen();
    pptx.layout = 'LAYOUT_16x9';

    for (const slideData of data.slides) {
      const slide = pptx.addSlide();
      
      // Judul Slide
      slide.addText(slideData.title, {
        x: 0.5, y: 0.5, w: '90%', h: 1,
        fontSize: 28, bold: true, color: '363636'
      });

      // Bullet Points
      const bulletText = slideData.bullets.map((b: string) => ({
        text: b,
        options: { bullet: true, fontSize: 18, color: '666666' }
      }));

      slide.addText(bulletText, {
        x: 0.5, y: 1.8, w: '90%', h: '70%',
        valign: 'top'
      });
    }

    // Konversi ke Buffer dan return sebagai response
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
