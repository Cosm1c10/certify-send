import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { CertificateData } from '@/types/certificate';

function calculateDaysToExpiry(expiryDate: string): number | null {
  if (!expiryDate || expiryDate === 'Not Found' || expiryDate === '') {
    return null;
  }

  const expiry = new Date(expiryDate);
  if (isNaN(expiry.getTime())) {
    return null;
  }

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  expiry.setHours(0, 0, 0, 0);

  const diffTime = expiry.getTime() - today.getTime();
  const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

  return diffDays;
}

// Client terminology for status
function getClientStatus(daysToExpiry: number | null): string {
  if (daysToExpiry === null) {
    return 'Unknown';
  }

  if (daysToExpiry < 0) {
    return 'Expired';
  } else if (daysToExpiry >= 0 && daysToExpiry < 30) {
    return 'Expiring Soon';
  } else {
    return 'Up to date';
  }
}

export const exportToExcel = async (certificates: CertificateData[]) => {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Certificate Analyzer';
  workbook.created = new Date();

  const worksheet = workbook.addWorksheet('Certificates');

  // Sort certificates alphabetically by supplier_name for grouping
  const sortedCertificates = [...certificates].sort((a, b) => {
    const nameA = (a.supplierName || '').toLowerCase();
    const nameB = (b.supplierName || '').toLowerCase();
    return nameA.localeCompare(nameB);
  });

  // Define columns with headers and widths (matching client's Master File)
  worksheet.columns = [
    { header: 'Supplier Name', key: 'supplierName', width: 30 },
    { header: 'Certificate / Report No', key: 'certificateNumber', width: 22 },
    { header: 'Country', key: 'country', width: 15 },
    { header: 'EC Regulation Measure', key: 'ecRegulation', width: 40 },
    { header: 'Certification', key: 'certification', width: 18 },
    { header: 'Issued', key: 'dateIssued', width: 12 },
    { header: 'Date of Expiry', key: 'dateExpired', width: 14 },
    { header: 'Status', key: 'certStatus', width: 14 },
    { header: 'Days to Expire', key: 'daysToExpiry', width: 14 },
  ];

  // Style the header row
  const headerRow = worksheet.getRow(1);
  headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  headerRow.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF4472C4' }, // Blue header
  };
  headerRow.alignment = { horizontal: 'center', vertical: 'middle' };
  headerRow.height = 25;

  // Add data rows (using sorted certificates for supplier grouping)
  sortedCertificates.forEach((cert) => {
    const daysToExpiry = calculateDaysToExpiry(cert.expiryDate);
    const status = getClientStatus(daysToExpiry);

    worksheet.addRow({
      supplierName: cert.supplierName || '',
      certificateNumber: cert.certificateNumber || '',
      country: cert.country || '',
      ecRegulation: cert.ecRegulation || '',
      certification: cert.certification || '',
      dateIssued: cert.issueDate || '',
      dateExpired: cert.expiryDate || '',
      certStatus: status,
      daysToExpiry: daysToExpiry !== null ? daysToExpiry : '',
    });
  });

  // Apply conditional formatting to data rows
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return; // Skip header

    const statusCell = row.getCell('certStatus');
    const daysCell = row.getCell('daysToExpiry');
    const status = statusCell.value as string;

    // Style the entire row based on status
    row.eachCell((cell) => {
      cell.border = {
        top: { style: 'thin', color: { argb: 'FFD9D9D9' } },
        left: { style: 'thin', color: { argb: 'FFD9D9D9' } },
        bottom: { style: 'thin', color: { argb: 'FFD9D9D9' } },
        right: { style: 'thin', color: { argb: 'FFD9D9D9' } },
      };
    });

    // Apply status-specific styling
    if (status === 'Expired') {
      // Red background for expired
      statusCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFFCCCC' }, // Light red
      };
      statusCell.font = { bold: true, color: { argb: 'FF9C0006' } }; // Dark red text

      daysCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFFCCCC' },
      };
      daysCell.font = { bold: true, color: { argb: 'FF9C0006' } };
    } else if (status === 'Expiring Soon') {
      // Orange background for expiring soon
      statusCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFFEB9C' }, // Light orange
      };
      statusCell.font = { bold: true, color: { argb: 'FF9C5700' } }; // Dark orange text

      daysCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFFEB9C' },
      };
      daysCell.font = { bold: true, color: { argb: 'FF9C5700' } };
    } else if (status === 'Up to date') {
      // Green background for valid
      statusCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFC6EFCE' }, // Light green
      };
      statusCell.font = { bold: true, color: { argb: 'FF006100' } }; // Dark green text

      daysCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFC6EFCE' },
      };
      daysCell.font = { bold: true, color: { argb: 'FF006100' } };
    }
  });

  // Add alternating row colors for better readability
  // Column indices: 1=Supplier, 2=CertNo, 3=Country, 4=ECReg, 5=Cert, 6=Issued, 7=Expiry, 8=Status, 9=Days
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    if (rowNumber % 2 === 0) {
      row.eachCell((cell, colNumber) => {
        // Don't override status/days cells (columns 8 and 9)
        if (colNumber !== 8 && colNumber !== 9) {
          if (!cell.fill || (cell.fill as ExcelJS.FillPattern).pattern !== 'solid') {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'FFF5F5F5' }, // Light gray for even rows
            };
          }
        }
      });
    }
  });

  // Freeze the header row
  worksheet.views = [{ state: 'frozen', ySplit: 1 }];

  // Generate buffer and save
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  });

  const date = new Date().toISOString().split('T')[0];
  saveAs(blob, `certificates-${date}.xlsx`);
};
