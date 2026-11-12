const en = {
  opening: 'Opening…',
  loadFailed: "Couldn't read this file.",
  pdfInvalid: "This file isn't a valid PDF, or it is damaged.",
  pdfPassword: "Password-protected PDFs can't be opened yet.",
  docxInvalid: "This file isn't a valid Word document, or it is damaged.",
  notYetTitle: "This file type can't be viewed yet",
  notYetBody: 'Viewing {type} files is coming in a later version. The file was read correctly and has not been changed.',
  notZip: "This doesn't look like a valid {type} file.",
  excel: 'Excel',
  powerpoint: 'PowerPoint',
  pages: 'Pages',
  search: 'Search',
  searchHint: 'Search in document',
  previous: 'Previous result',
  next: 'Next result',
  close: 'Close',
  noResults: 'No results',
  page: 'Page {n}',
  chooseFile: 'Choose a file',
};

const id: typeof en = {
  opening: 'Membuka…',
  loadFailed: 'File ini tidak bisa dibaca.',
  pdfInvalid: 'File ini bukan PDF yang valid, atau rusak.',
  pdfPassword: 'PDF yang dilindungi kata sandi belum bisa dibuka.',
  docxInvalid: 'File ini bukan dokumen Word yang valid, atau rusak.',
  notYetTitle: 'Jenis file ini belum bisa ditampilkan',
  notYetBody: 'Tampilan file {type} akan hadir di versi berikutnya. File terbaca dengan benar dan tidak diubah.',
  notZip: 'Ini sepertinya bukan file {type} yang valid.',
  excel: 'Excel',
  powerpoint: 'PowerPoint',
  pages: 'Halaman',
  search: 'Cari',
  searchHint: 'Cari di dokumen',
  previous: 'Hasil sebelumnya',
  next: 'Hasil berikutnya',
  close: 'Tutup',
  noResults: 'Tidak ada hasil',
  page: 'Halaman {n}',
  chooseFile: 'Pilih file',
};

const strings = /^(id|in)\b/i.test(navigator.language) ? id : en;

export function t(key: keyof typeof en, params: Record<string, string | number> = {}): string {
  return strings[key].replace(/\{(\w+)\}/g, (_, name: string) => String(params[name] ?? ''));
}
