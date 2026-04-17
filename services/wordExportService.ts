import { saveAs } from 'file-saver';
import { jsPDF } from 'jspdf';
import { toPng } from 'html-to-image';

export interface ExportMetadata {
  author?: string;
  date?: string;
  title?: string;
}

export const exportToWord = (
  htmlContent: string, 
  filename: string, 
  headerHtml: string = '', 
  marginValue: string = '0.4in 0.6in 0.4in 0.6in',
  fontFamily: string = 'Times New Roman',
  lineHeight: string = '1.15',
  metadata?: ExportMetadata,
  isFrameEnabled: boolean = false,
  activeDesign: string = '',
  paperStyles?: any,
  mcqStyle: number = 0,
  globalLayout: number = 0,
  baseLayout: number = 0,
  instructionRulerStyle: number = 0,
  instructionHeaderStyle: number = 0,
  instructionStyle: number = 0,
  isInstructionBackgroundEnabled: boolean = false,
  isColorExportEnabled: boolean = false,
  exportTheme: number = 1
) => {
  const tempDiv = document.createElement('div');
  tempDiv.innerHTML = htmlContent;

  const headerDiv = document.createElement('div');
  headerDiv.innerHTML = headerHtml;

  // Randomize Ruler Color if it's the middle ruler layout
  let activeRulerColor = '#334155'; // Use a neutral slate instead of purple
  if (globalLayout === 1 || activeDesign === 'design-playful') activeRulerColor = '#059669'; // Green for Orange Mix left bar
  else if (globalLayout === 2 || activeDesign === 'design-eco') activeRulerColor = '#059669';
  else if (activeDesign === 'design-modern-blue') activeRulerColor = '#2563eb';
  else if (globalLayout === 3) activeRulerColor = '#9333ea'; // Purple for Soft Lavender

  // Dynamic Line Spacing Logic
  const spacingMap: Record<string, string> = {
    '1.0': '15pt',
    '1.15': '18pt',
    '1.5': '24pt',
    '2.0': '32pt'
  };
  const exactLineHeight = spacingMap[lineHeight] || `${Math.round(parseFloat(lineHeight) * 16)}pt`;

  // 1. Image Formatting (Synchronous to prevent Chrome's strict gesture-expiry download interruptions)
  const images = [...Array.from(tempDiv.querySelectorAll('img')), ...Array.from(headerDiv.querySelectorAll('img'))];
  for (const img of images) {
    const originalWidth = img.width || 550;
    const isLogo = img.style.maxHeight === '80pt' || img.classList.contains('logo') || headerDiv.contains(img);

    if (isLogo) {
      img.style.width = '1.25in';
      img.style.height = 'auto';
    } else if (originalWidth > 200) {
      img.style.width = '6.5in';
      img.style.height = 'auto';
    } else {
      img.style.width = `${(originalWidth / 96).toFixed(2)}in`;
      img.style.height = 'auto';
    }
    img.style.display = 'block';
    if (!isLogo) img.style.margin = '5px auto';
  }

  // MCQ Styling
  const designClass = activeDesign || '';

  // Comprehensive MCQ Styles
  const mcqElements = tempDiv.querySelectorAll('b, strong, span');
  mcqElements.forEach(el => {
    if (mcqStyle > 0) {
      let text = el.textContent?.trim().toUpperCase() || '';
      text = text.replace(/[\(\)\[\]\.\s]/g, '');
      if (['A', 'B', 'C', 'D'].includes(text) && text.length === 1) {
        // Base styling for all MCQs
        if (mcqStyle !== 1 && mcqStyle !== 15) {
          (el as HTMLElement).style.display = 'inline-block';
          (el as HTMLElement).style.width = '22pt';
          (el as HTMLElement).style.height = '22pt';
          (el as HTMLElement).style.lineHeight = '22pt';
          (el as HTMLElement).style.textAlign = 'center';
          (el as HTMLElement).style.marginRight = '6pt';
          (el as HTMLElement).style.fontWeight = 'bold';
          (el as HTMLElement).style.fontSize = '10pt';
          (el as HTMLElement).style.verticalAlign = 'middle';
          
          if (designClass === 'design-modern-blue') {
            (el as HTMLElement).style.border = '1.5pt solid #2563eb';
            (el as HTMLElement).style.backgroundColor = '#eff6ff';
            (el as HTMLElement).style.borderRadius = '11pt';
          } else if (designClass === 'design-playful') {
            (el as HTMLElement).style.border = '2pt solid #f97316';
            (el as HTMLElement).style.backgroundColor = '#ffedd5';
            (el as HTMLElement).style.borderRadius = '11pt';
          } else {
            (el as HTMLElement).style.border = '1pt solid black';
            if (mcqStyle === 3) (el as HTMLElement).style.borderRadius = '11pt';
          }
        }

        let borderColor = 'black';
        let bgColor = 'transparent';
        let textColor = 'black';
        let isFilled = 'f';
        let strokeWt = '1pt';

        if (isColorExportEnabled) {
          if (designClass === 'design-modern-blue' && (mcqStyle === 1 || mcqStyle === 15)) {
            borderColor = '#2563eb';
            bgColor = '#eff6ff';
            textColor = '#2563eb';
            isFilled = 't';
            strokeWt = '1.5pt';
          } else if (designClass === 'design-playful' && (mcqStyle === 1 || mcqStyle === 15)) {
            borderColor = '#f97316';
            bgColor = '#ffedd5';
            textColor = '#ea580c';
            isFilled = 't';
            strokeWt = '1.5pt';
          }
        }

        if (mcqStyle === 1 || mcqStyle === 15) {
          // Both Round (1) and Crocodile Egg (15) use VML to ensure circles stay circles in Word.
          // Crocodile Egg just happens to trigger the playful/modern colors naturally.
          if (mcqStyle === 15 && isColorExportEnabled) {
            borderColor = '#059669';
            bgColor = '#ecfdf5';
            textColor = '#059669';
            isFilled = 't';
            strokeWt = '1.5pt';
          }
          let vmlFill = isFilled === 't' ? `fillcolor="${bgColor}"` : `filled="f"`;
          let htmlBg = isFilled === 't' ? `background:${bgColor};` : 'background:transparent;';
          let htmlBorder = isFilled === 't' ? `1.5pt solid ${borderColor}` : `1pt solid black`;
          
          el.innerHTML = `<!--[if gte vml 1]><v:oval style="width:18pt;height:18pt;position:relative;" ${vmlFill} strokecolor="${borderColor}" strokeweight="${strokeWt}"><v:textbox inset="0,0,0,0" style="mso-fit-shape-to-text:true;"><div style="text-align:center;font-size:10pt;color:${textColor};font-weight:bold;margin-top:0pt;margin-left:0pt;">${text}</div></v:textbox></v:oval><![endif]--><!--[if !mso]>--><span style="border:${htmlBorder}; padding:1pt 4pt; border-radius:7.2pt; ${htmlBg} color:${textColor}; font-weight:bold; font-size:12pt;">${text}</span><!--<![endif]-->`;
          (el as HTMLElement).style.border = 'none';
          (el as HTMLElement).style.setProperty('mso-border-alt', 'none');
          (el as HTMLElement).style.padding = '0';
          (el as HTMLElement).style.backgroundColor = 'transparent';
          (el as HTMLElement).style.setProperty('mso-shading', 'transparent');
          (el as HTMLElement).style.display = 'inline-block';
          (el as HTMLElement).style.textAlign = 'center';
          (el as HTMLElement).style.fontWeight = 'bold';
          (el as HTMLElement).style.marginRight = '6pt';
          (el as HTMLElement).style.verticalAlign = 'middle';
        }
        else if (mcqStyle === 2) el.innerHTML = `[${text}]`;
        else if (mcqStyle === 6) el.innerHTML = `◆${text}`;
        else if (mcqStyle === 8) {
          el.innerHTML = text === 'A' ? 'Ⓐ' : text === 'B' ? 'Ⓑ' : text === 'C' ? 'Ⓒ' : 'Ⓓ';
          (el as HTMLElement).style.border = 'none';
          (el as HTMLElement).style.setProperty('mso-border-alt', 'none');
          (el as HTMLElement).style.backgroundColor = 'transparent';
          (el as HTMLElement).style.setProperty('mso-shading', 'transparent');
        }
        else if (mcqStyle === 11 || mcqStyle === 12) {
          // Double Circle / Dotted Circle -> use Unicode circle 
          el.innerHTML = `<span style="font-size:12pt;">${text === 'A' ? 'Ⓐ' : text === 'B' ? 'Ⓑ' : text === 'C' ? 'Ⓒ' : 'Ⓓ'}</span>`;
          (el as HTMLElement).style.border = 'none';
          (el as HTMLElement).style.setProperty('mso-border-alt', 'none');
          (el as HTMLElement).style.backgroundColor = 'transparent';
          (el as HTMLElement).style.setProperty('mso-shading', 'transparent');
        }
        else if (mcqStyle === 13 || mcqStyle === 14) {
          // Solid background circles -> use dark Unicode to avoid square background
          el.innerHTML = text === 'A' ? '🅐' : text === 'B' ? '🅑' : text === 'C' ? '🅒' : '🅓';
          (el as HTMLElement).style.border = 'none';
          (el as HTMLElement).style.setProperty('mso-border-alt', 'none');
          (el as HTMLElement).style.backgroundColor = 'transparent';
          (el as HTMLElement).style.setProperty('mso-shading', 'transparent');
          if (mcqStyle === 14) (el as HTMLElement).style.color = '#10b981'; // Playful green
        }
        
        if ([8, 11, 12, 13, 14].includes(mcqStyle)) {
          (el as HTMLElement).style.borderRadius = '11pt';
          (el as HTMLElement).style.textAlign = 'center';
        }
      }
    }
  });

  // Instruction Rulers
  const existingRulers = tempDiv.querySelectorAll('[class*="instruction-ruler-"]');
  existingRulers.forEach(ruler => {
    const el = ruler as HTMLElement;
    const styleNum = parseInt(el.className.match(/instruction-ruler-(\d+)/)?.[1] || '0');
    el.style.width = '100%';
    el.style.margin = '5pt 0 10pt 0';
    el.innerHTML = '&nbsp;';
    if (styleNum === 1) el.style.borderBottom = `1.5pt solid ${activeRulerColor}`;
    else if (styleNum === 2) el.style.borderBottom = `2pt dashed ${activeRulerColor}`;
    else if (styleNum === 3) el.style.borderBottom = `4pt double ${activeRulerColor}`;
    else if (styleNum === 4) el.style.borderBottom = `4pt solid ${activeRulerColor}`;
    
    if (styleNum > 0) el.style.setProperty('mso-border-bottom-alt', el.style.borderBottom);
  });

  // Table Logic
  const optionsTables = tempDiv.querySelectorAll('.options-table, [data-type="mcq-options"]');
  optionsTables.forEach(table => {
    (table as HTMLElement).style.border = 'none';
    (table as HTMLElement).style.setProperty('mso-border-alt', 'none');
    table.querySelectorAll('td').forEach(cell => {
      (cell as HTMLElement).style.border = 'none';
      (cell as HTMLElement).style.setProperty('mso-border-alt', 'none');
      (cell as HTMLElement).style.padding = '4pt 8pt';
    });
  });

  const tables = tempDiv.querySelectorAll('table');
  tables.forEach(table => {
    const isNested = table.parentElement?.closest('table') !== null;
    if (isNested) {
      table.style.border = 'none';
      table.style.width = '100%';
      table.querySelectorAll('td').forEach(c => {
        (c as HTMLElement).style.border = 'none';
        (c as HTMLElement).style.padding = '2pt';
      });
    } else {
      const isRulerTable = table.classList.contains('ruler-table') || table.rows[0]?.cells.length === 2;
      if (isRulerTable) {
        table.style.border = 'none';
        table.style.borderCollapse = 'collapse';
        Array.from(table.rows).forEach(row => {
          Array.from(row.cells).forEach((c, idx) => {
            const cell = c as HTMLElement;
            cell.style.padding = '15pt';
            cell.style.border = 'none';
            if (idx === 0 && row.cells.length === 2) {
              cell.style.borderRight = `1.5pt solid ${activeRulerColor}`;
              cell.style.setProperty('mso-border-right-alt', `1.5pt solid ${activeRulerColor}`);
            }
          });
        });
      }
      
      // Header Styles - STRICTLY respect isInstructionBackgroundEnabled
      table.querySelectorAll('td').forEach(cell => {
        const isHeader = cell.classList.contains('header-row') || (cell.getAttribute('colspan') === '2');
        if (isHeader) {
          const c = cell as HTMLElement;
          // FORCE white background if disabled to prevent Word inheritance/defaults
          let bg = '#ffffff';
          let textColor = '#000000';
          let shading = '#ffffff';

          if (isInstructionBackgroundEnabled) {
            // Apply specific style mappings based on the chosen style
            const headerStyles: Record<number, { bg: string, color: string, border?: string }> = {
              0: { bg: '#facc15', color: '#000000', border: '3pt solid black' },
              1: { bg: '#f59e0b', color: '#ffffff' },
              3: { bg: '#1e293b', color: '#ffffff', border: '8pt solid #6366f1' },
              4: { bg: '#dcfce7', color: '#065f46', border: '2pt solid #10b981' },
              5: { bg: '#fde047', color: '#000000', border: '4pt solid black' },
              6: { bg: '#4f46e5', color: '#ffffff' },
              13: { bg: '#dcfce7', color: '#064e3b', border: '3pt solid #059669' },
              15: { bg: '#581c87', color: '#ffffff', border: '2pt solid #fbbf24' },
              19: { bg: '#ea580c', color: '#ffffff' }
            };

            const style = headerStyles[instructionHeaderStyle] || { bg: '#dcfce7', color: '#064e3b' };
            bg = style.bg;
            textColor = style.color;
            shading = style.bg;
            if (style.border) {
              c.style.border = style.border;
              c.style.setProperty('mso-border-alt', style.border);
            }
          }

          c.style.backgroundColor = bg;
          c.style.setProperty('mso-shading', shading);
          c.style.color = textColor;
          
          if (!isInstructionBackgroundEnabled) {
            c.style.border = 'none';
            c.style.borderBottom = `1.5pt solid ${activeRulerColor}`;
            c.style.setProperty('mso-border-bottom-alt', `1.5pt solid ${activeRulerColor}`);
          } else if (!c.style.border) {
            c.style.borderLeft = `6pt solid ${activeRulerColor}`;
            c.style.setProperty('mso-border-left-alt', `6pt solid ${activeRulerColor}`);
          }
          c.style.padding = '10pt';
          c.style.paddingLeft = '15pt';
          c.style.fontWeight = 'bold';
        }
      });
      
      // Zebra Striping detection
      const rows = Array.from(table.rows);
      if (table.classList.contains('zebra') || table.getAttribute('data-type') === 'zebra') {
        rows.forEach((row, idx) => {
          if (idx % 2 === 1) { // odd index = even row (1, 3, 5...)
            Array.from(row.cells).forEach(cell => {
              (cell as HTMLElement).style.backgroundColor = '#f8fafc';
              (cell as HTMLElement).style.setProperty('mso-shading', '#f8fafc');
            });
          }
        });
      }
    }
  });

  // Word Bank Box
  const wordBanks = tempDiv.querySelectorAll('.word-bank-box-alt, .word-bank');
  wordBanks.forEach(box => {
    const el = box as HTMLElement;
    el.style.border = '1.5pt solid #334155';
    el.style.padding = '10pt';
    el.style.margin = '10pt 0';
    el.style.backgroundColor = '#f1f5f9';
    el.style.setProperty('mso-shading', '#f1f5f9');
    el.style.textAlign = 'center';
    el.style.fontWeight = 'bold';
    el.style.borderRadius = '5pt';
  });

  // Unwrap prose
  let sections = Array.from(tempDiv.children);
  if (sections.length === 1 && sections[0].classList.contains('prose')) {
    sections = Array.from(sections[0].children);
  }

  let finalHtml = "";
  sections.forEach(el => {
    finalHtml += `
      <table border="0" cellspacing="0" cellpadding="0" width="100%" style="width: 100%; border-collapse: collapse;">
        <tr>
          <td style="font-family: '${fontFamily}', serif; font-size: 12pt; line-height: ${exactLineHeight}; mso-line-height-rule: exactly;">
            ${(el as HTMLElement).outerHTML}
          </td>
        </tr>
      </table>`;
  });

  // 5. Frame Style Simulation
  const frameStyle = isFrameEnabled ? `border: 2pt solid ${activeRulerColor}; mso-border-alt: 2.5pt solid ${activeRulerColor}; border-radius: 15pt; mso-border-radius: 15pt; padding: 15pt;` : '';

  // Paper Styles - Moved All Borders to TD level for better Word support
  let bodyBgColor = '#ffffff';
  let paperTdStyle = '';
  
  if (globalLayout === 1) { // Orange Mix (Green & Orange feel)
    paperTdStyle = `border-left: 15pt solid #059669; mso-border-left-alt: 15pt solid #059669; border-top: 15pt solid #ea580c; mso-border-top-alt: 15pt solid #ea580c; border-top-left-radius: 40pt; padding-left: 20pt; padding-top: 20pt; background: #ffffff; mso-shading: windowtext 0% #ffffff;`;
  } else if (globalLayout === 2) { // Modern Emerald
    paperTdStyle = `background-color: #f0fdf4; border-left: 15pt solid #059669; padding-left: 15pt; mso-shading: windowtext 0% #f0fdf4;`;
    bodyBgColor = '#f0fdf4';
  } else if (globalLayout === 17) {
    paperTdStyle = `background-color: #ffffff; border-left: 4.5pt double #ef4444; padding-left: 35pt; mso-shading: windowtext 0% #ffffff;`;
  } else if (globalLayout === 18) {
    paperTdStyle = `background-color: #fef3c7; border: 1pt solid #fde68a; mso-shading: windowtext 0% #fef3c7;`;
    bodyBgColor = '#fef3c7';
  } else {
    paperTdStyle = `mso-shading: windowtext 0% ${bodyBgColor};`;
  }

  const content = `
    <html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
    <head><meta charset='utf-8'>
      <style>
        @page Section1 { size: 8.5in 11.0in; margin: 0.5in; }
        div.Section1 { page: Section1; }
        body { font-family: "${fontFamily}", serif; font-size: 12pt; line-height: ${exactLineHeight}; mso-line-height-rule: exactly; background-color: ${bodyBgColor}; }
        table { border-collapse: collapse; width: 100%; }
        td { padding: 0; vertical-align: top; }
        .options-table td { mso-line-height-rule: at-least; line-height: 24pt; height: 26pt; }
      </style>
    </head>
    <body>
      <div class="Section1">
        <!-- Master Table for Paper Design -->
        <table border="0" cellspacing="0" cellpadding="0" width="100%" style="width: 100%; border-collapse: collapse;">
          <tr>
            <td style="padding: 30pt; ${paperTdStyle} ${frameStyle}">
              <div style="${globalLayout === 1 ? 'border-left: 3pt solid #059669; padding-left: 15pt; mso-border-left-alt: 3.5pt solid #059669;' : ''}">
                ${headerDiv.innerHTML}
                ${finalHtml}
              </div>
            </td>
          </tr>
        </table>
      </div>
    </body>
    </html>`;

  const blob = new Blob(['\ufeff', content], { type: 'application/msword' });
  saveAs(blob, `${filename}.doc`);
};

export const exportToHTML = (htmlContent: string, filename: string, headerHtml: string = '') => {
  const fullHtml = `<html><body><div class="header">${headerHtml}</div><div class="content">${htmlContent}</div></body></html>`;
  saveAs(new Blob([fullHtml], { type: 'text/html;charset=utf-8' }), `${filename}.html`);
};

export const exportToPDF = async (elementId: string, filename: string) => {
  const element = document.getElementById(elementId);
  if (!element) return;
  try {
    // High-resolution capture (300 DPI approx)
    const dataUrl = await toPng(element, { 
      quality: 1,
      pixelRatio: 2, // Double pixels for crispness
      skipFonts: false,
      cacheBust: true
    });
    
    const pdf = new jsPDF('p', 'mm', 'a4');
    const pdfWidth = pdf.internal.pageSize.getWidth();
    const pdfHeight = pdf.internal.pageSize.getHeight();
    
    // Calculate dimensions to fit the page
    const imgWidth = pdfWidth;
    const imgHeight = (element.offsetHeight * pdfWidth) / element.offsetWidth;
    
    // Handle multi-page if content is too long
    let heightLeft = imgHeight;
    let position = 0;
    
    pdf.addImage(dataUrl, 'PNG', 0, position, imgWidth, imgHeight, undefined, 'FAST');
    heightLeft -= pdfHeight;
    
    while (heightLeft >= 0) {
      position = heightLeft - imgHeight;
      pdf.addPage();
      pdf.addImage(dataUrl, 'PNG', 0, position, imgWidth, imgHeight, undefined, 'FAST');
      heightLeft -= pdfHeight;
    }
    
    pdf.save(`${filename}.pdf`);
  } catch (error) {
    console.error("PDF Export failed", error);
    window.print();
  }
};
