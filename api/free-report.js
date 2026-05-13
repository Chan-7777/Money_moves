const { generateReport } = require('../build_pdf_report.js');

async function convertToPdf(docxBuffer) {
  const form = new FormData();
  form.append(
    'File',
    new Blob([docxBuffer], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }),
    'report.docx'
  );
  const res = await fetch('https://v2.convertapi.com/convert/docx/to/pdf', {
    method: 'POST',
    headers: { Authorization: `Bearer ${process.env.CONVERTAPI_SECRET}` },
    body: form,
  });
  if (!res.ok) throw new Error(`ConvertAPI error: ${res.status}`);
  const data = await res.json();
  return Buffer.from(data.Files[0].FileData, 'base64');
}

module.exports = async (req, res) => {
  if (req.method !== 'POST') return res.status(405).end();

  const { state } = req.body || {};
  if (!state) return res.status(400).json({ error: 'Missing state' });

  try {
    const docxBuffer = await generateReport(state);
    const pdfBuffer  = await convertToPdf(docxBuffer);

    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', 'attachment; filename="MoneyMoves_AU_Plan.pdf"');
    res.setHeader('Content-Length', pdfBuffer.length);
    res.status(200).send(pdfBuffer);
  } catch (err) {
    console.error('[free-report]', err);
    res.status(500).json({ error: err.message });
  }
};
