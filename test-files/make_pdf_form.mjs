// Writes test-files/pdf-form.pdf: a fillable form (text, multi-line text, checkbox, radio group,
// dropdown) and a second page turned with /Rotate 90, for testing PDF editing.
//
//     node test-files/make_pdf_form.mjs      (uses web/node_modules/pdf-lib)

import { writeFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const require = createRequire(join(here, '../web/package.json'));
const { PDFDocument, StandardFonts, rgb, degrees } = require('pdf-lib');

const doc = await PDFDocument.create();
const font = await doc.embedFont(StandardFonts.Helvetica);
const bold = await doc.embedFont(StandardFonts.HelveticaBold);
const page = doc.addPage([595, 842]);
const form = doc.getForm();

page.drawText('Membership form', { x: 56, y: 770, size: 24, font: bold });
page.drawText('Fill in the fields below. This paragraph is here to be highlighted and underlined', { x: 56, y: 735, size: 12, font });
page.drawText('while testing; it runs over two lines so a selection can span both of them.', { x: 56, y: 719, size: 12, font });

const label = (text, y) => page.drawText(text, { x: 56, y, size: 12, font });
label('Full name', 676);
const name = form.createTextField('name');
name.addToPage(page, { x: 180, y: 668, width: 300, height: 22 });

label('Email', 636);
const email = form.createTextField('email');
email.setText('someone@example.com');
email.addToPage(page, { x: 180, y: 628, width: 300, height: 22 });

label('Notes', 596);
const notes = form.createTextField('notes');
notes.enableMultiline();
notes.addToPage(page, { x: 180, y: 530, width: 300, height: 80 });

label('Newsletter', 496);
const news = form.createCheckBox('newsletter');
news.addToPage(page, { x: 180, y: 492, width: 16, height: 16 });

label('Plan', 456);
const plan = form.createRadioGroup('plan');
for (const [i, option] of ['basic', 'plus', 'pro'].entries()) {
  plan.addOptionToPage(option, page, { x: 180 + i * 90, y: 452, width: 16, height: 16 });
  page.drawText(option, { x: 200 + i * 90, y: 456, size: 12, font });
}
plan.select('basic');

label('Country', 416);
const country = form.createDropdown('country');
country.addOptions(['Indonesia', 'Malaysia', 'Singapore', 'Other']);
country.select('Indonesia');
country.addToPage(page, { x: 180, y: 408, width: 200, height: 22 });

page.drawText('Signature:', { x: 56, y: 330, size: 12, font });
page.drawLine({ start: { x: 130, y: 326 }, end: { x: 400, y: 326 }, thickness: 1, color: rgb(0.4, 0.4, 0.4) });

const turned = doc.addPage([595, 842]);
turned.setRotation(degrees(90));
turned.drawText('Page 2 is stored upright but shown turned 90 degrees (/Rotate 90).', { x: 56, y: 770, size: 14, font });
turned.drawText('Signatures and drawings placed on it must still look upright on screen.', { x: 56, y: 750, size: 12, font });

writeFileSync(join(here, 'pdf-form.pdf'), await doc.save());
console.log('wrote test-files/pdf-form.pdf');
