const express = require('express');
const axios = require('axios');
const cheerio = require('cheerio');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType, LevelFormat
} = require('docx');

const app = express();
app.use(express.json({ limit: '10mb' }));

// ─────────────────────────────────────────
// CORS — 브라우저에서의 직접 호출 허용
// ─────────────────────────────────────────
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.header('Access-Control-Allow-Headers', 'Content-Type, Authorization');
  if (req.method === 'OPTIONS') return res.sendStatus(200);
  next();
});

// ─────────────────────────────────────────
// 헬스체크 (Render가 서버 상태 확인용으로 사용)
// ─────────────────────────────────────────
app.get('/', (req, res) => {
  res.json({ status: 'ok', message: '회의록 생성 서버 가동 중' });
});

// ─────────────────────────────────────────
// 기사 본문 크롤링 엔드포인트
// POST /fetch-article
// Body: { url: string }
// 응답: { text: string }
// ─────────────────────────────────────────
app.post('/fetch-article', async (req, res) => {
  const { url } = req.body;
  if (!url) return res.status(400).json({ text: '' });

  try {
    const response = await axios.get(url, {
      timeout: 8000,
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'ko-KR,ko;q=0.9,en;q=0.8',
      },
      maxRedirects: 5,
    });

    const $ = cheerio.load(response.data);

    // 노이즈 제거
    $('script, style, nav, header, footer, aside, .ad, .ads, .advertisement, .cookie, .popup, .modal, .sidebar, .related, .comment, .share, noscript, iframe').remove();

    // 본문 추출 (우선순위 순)
    const selectors = [
      'article',
      '[role="main"]',
      'main',
      '.article-body',
      '.article-content',
      '.post-body',
      '.post-content',
      '.entry-content',
      '.content-body',
      '.story-body',
      '#article-body',
      '#content',
      '.content',
    ];

    let text = '';
    for (const sel of selectors) {
      const el = $(sel).first();
      if (el.length) {
        text = el.text();
        break;
      }
    }

    // fallback: body 전체
    if (!text.trim()) {
      text = $('body').text();
    }

    // 공백 정리 및 3000자 truncate
    text = text
      .replace(/\s+/g, ' ')
      .replace(/\n{3,}/g, '\n\n')
      .trim()
      .slice(0, 3000);

    res.json({ text });
  } catch (err) {
    console.error(`fetch-article 실패 [${url}]:`, err.message);
    res.json({ text: '' }); // 실패해도 에러 throw 안 함
  }
});

// ─────────────────────────────────────────
// 회의록 생성 엔드포인트
// POST /generate-minutes
// Body: { title, date, attendees, content }
// ─────────────────────────────────────────
app.post('/generate-minutes', async (req, res) => {
  try {
    const { title, date, attendees, content, webhook_secret } = req.body;

    // 간단한 보안 체크 (환경변수로 설정)
    if (process.env.WEBHOOK_SECRET && webhook_secret !== process.env.WEBHOOK_SECRET) {
      return res.status(401).json({ error: 'Unauthorized' });
    }

    if (!content) {
      return res.status(400).json({ error: 'content 필드가 필요합니다.' });
    }

    const docTitle = title || '회의록';
    const docDate = date || new Date().toLocaleDateString('ko-KR');
    const docAttendees = attendees || '';

    // ── 문서 스타일 헬퍼 ──
    const border = { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' };
    const borders = { top: border, bottom: border, left: border, right: border };

    const divider = new Paragraph({
      border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC', space: 1 } },
      spacing: { after: 160 },
      children: []
    });

    function bodyText(text) {
      return new Paragraph({
        spacing: { after: 140 },
        children: [new TextRun({ text, size: 22, font: '맑은 고딕' })]
      });
    }

    // Claude가 반환한 content를 섹션별로 파싱해서 문단 생성
    const contentParagraphs = parseContentToParagraphs(content);

    const doc = new Document({
      styles: {
        default: { document: { run: { font: '맑은 고딕', size: 22 } } },
        paragraphStyles: [
          {
            id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
            run: { size: 28, bold: true, font: '맑은 고딕', color: '1F3864' },
            paragraph: { spacing: { before: 320, after: 120 }, outlineLevel: 0 }
          },
          {
            id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
            run: { size: 24, bold: true, font: '맑은 고딕', color: '2E4F8A' },
            paragraph: { spacing: { before: 240, after: 80 }, outlineLevel: 1 }
          }
        ]
      },
      sections: [{
        properties: {
          page: {
            size: { width: 11906, height: 16838 },
            margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
          }
        },
        children: [
          // 제목
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 200 },
            children: [new TextRun({ text: '회  의  록', bold: true, size: 40, font: '맑은 고딕', color: '1F3864' })]
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 400 },
            children: [new TextRun({ text: docTitle, size: 26, font: '맑은 고딕', color: '555555' })]
          }),

          // 기본 정보 테이블
          new Table({
            width: { size: 9026, type: WidthType.DXA },
            columnWidths: [2000, 7026],
            rows: [
              makeInfoRow('일시', docDate, borders),
              makeInfoRow('참석자', docAttendees, borders),
            ]
          }),

          new Paragraph({ spacing: { after: 300 }, children: [] }),
          divider,

          // Claude가 정리한 본문 내용
          ...contentParagraphs,
        ]
      }]
    });

    const buffer = await Packer.toBuffer(doc);

    res.set({
      'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'Content-Disposition': `attachment; filename="${encodeURIComponent(docTitle)}.docx"`,
      'Content-Length': buffer.length
    });
    res.send(buffer);

  } catch (err) {
    console.error('문서 생성 오류:', err);
    res.status(500).json({ error: err.message });
  }
});

// ─────────────────────────────────────────
// 헬퍼: 정보 테이블 행 생성
// ─────────────────────────────────────────
function makeInfoRow(label, value, borders) {
  return new TableRow({
    children: [
      new TableCell({
        borders,
        width: { size: 2000, type: WidthType.DXA },
        shading: { fill: 'EEF2F8', type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({ children: [new TextRun({ text: label, bold: true, size: 22, font: '맑은 고딕' })] })]
      }),
      new TableCell({
        borders,
        width: { size: 7026, type: WidthType.DXA },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({ children: [new TextRun({ text: value, size: 22, font: '맑은 고딕' })] })]
      })
    ]
  });
}

// ─────────────────────────────────────────
// 헬퍼: Claude 텍스트 → docx 문단 배열로 변환
// ## → Heading1, ### → Heading2, 나머지 → body
// ─────────────────────────────────────────
function parseContentToParagraphs(text) {
  const lines = text.split('\n');
  const paragraphs = [];

  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed) {
      paragraphs.push(new Paragraph({ spacing: { after: 80 }, children: [] }));
      continue;
    }

    if (trimmed.startsWith('## ')) {
      paragraphs.push(new Paragraph({
        heading: HeadingLevel.HEADING_1,
        spacing: { before: 320, after: 120 },
        children: [new TextRun({ text: trimmed.replace('## ', ''), bold: true, size: 28, font: '맑은 고딕', color: '1F3864' })]
      }));
    } else if (trimmed.startsWith('### ')) {
      paragraphs.push(new Paragraph({
        heading: HeadingLevel.HEADING_2,
        spacing: { before: 240, after: 80 },
        children: [new TextRun({ text: trimmed.replace('### ', ''), bold: true, size: 24, font: '맑은 고딕', color: '2E4F8A' })]
      }));
    } else if (trimmed.startsWith('**') && trimmed.endsWith('**')) {
      paragraphs.push(new Paragraph({
        spacing: { before: 160, after: 80 },
        children: [new TextRun({ text: trimmed.replace(/\*\*/g, ''), bold: true, size: 22, font: '맑은 고딕' })]
      }));
    } else {
      // 인라인 볼드(**text**) 처리
      const runs = parseInlineBold(trimmed);
      paragraphs.push(new Paragraph({
        spacing: { after: 140 },
        children: runs
      }));
    }
  }

  return paragraphs;
}

// ─────────────────────────────────────────
// 헬퍼: 인라인 **bold** 파싱
// ─────────────────────────────────────────
function parseInlineBold(text) {
  const parts = text.split(/(\*\*[^*]+\*\*)/g);
  return parts.map(part => {
    if (part.startsWith('**') && part.endsWith('**')) {
      return new TextRun({ text: part.slice(2, -2), bold: true, size: 22, font: '맑은 고딕' });
    }
    return new TextRun({ text: part, size: 22, font: '맑은 고딕' });
  });
}

// ─────────────────────────────────────────
// RSS XML 프록시 엔드포인트 (CORS 우회용)
// POST /fetch-rss
// Body: { url: string }
// 응답: { xml: string }
// ─────────────────────────────────────────
app.post('/fetch-rss', async (req, res) => {
  const { url } = req.body;
  if (!url) return res.status(400).json({ error: 'url 필드 필요' });

  try {
    const response = await axios.get(url, {
      timeout: 10000,
      headers: {
        'User-Agent': 'Mozilla/5.0 (compatible; RSS-Reader/1.0)',
        'Accept': 'application/rss+xml, application/xml, text/xml, */*',
      },
      maxRedirects: 5,
      responseType: 'text',
    });

    res.json({ xml: response.data });
  } catch (err) {
    console.error(`fetch-rss 실패 [${url}]:`, err.message);
    res.status(502).json({ error: err.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`서버 실행 중: http://localhost:${PORT}`);
});
