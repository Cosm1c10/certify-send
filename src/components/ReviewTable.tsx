import { CertificateData } from '@/types/certificate';
import {
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
} from '@/components/ui/table';
import { Badge } from '@/components/ui/badge';
import { FileText } from 'lucide-react';

interface ReviewTableProps {
  certificates: CertificateData[];
}

// Calculate days to expiry for status display
function getDaysToExpiry(expiryDate: string): number | null {
  if (!expiryDate || expiryDate === 'Not Found' || expiryDate === '') {
    return null;
  }
  const expiry = new Date(expiryDate);
  if (isNaN(expiry.getTime())) return null;
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  expiry.setHours(0, 0, 0, 0);
  return Math.ceil((expiry.getTime() - today.getTime()) / (1000 * 60 * 60 * 24));
}

// Get display status based on days to expiry
function getCertStatus(daysToExpiry: number | null): 'expired' | 'expiring_soon' | 'up_to_date' | 'unknown' {
  if (daysToExpiry === null) return 'unknown';
  if (daysToExpiry < 0) return 'expired';
  if (daysToExpiry < 30) return 'expiring_soon';
  return 'up_to_date';
}

const statusStyles = {
  expired: 'bg-red-500/15 text-red-700 border-red-500/30 font-medium',
  expiring_soon: 'bg-orange-500/15 text-orange-700 border-orange-500/30 font-medium',
  up_to_date: 'bg-green-500/15 text-green-700 border-green-500/30 font-medium',
  unknown: 'bg-muted text-muted-foreground border-border',
};

const statusLabels = {
  expired: 'Expired',
  expiring_soon: 'Expiring Soon',
  up_to_date: 'Up to date',
  unknown: 'Unknown',
};

const ReviewTable = ({ certificates }: ReviewTableProps) => {
  if (certificates.length === 0) {
    return (
      <div className="border border-border rounded-xl p-12 text-center">
        <FileText className="w-12 h-12 text-muted-foreground mx-auto mb-4" />
        <p className="text-muted-foreground">
          No certificates analyzed yet. Drop a PDF above to get started.
        </p>
      </div>
    );
  }

  return (
    <div className="border border-border rounded-xl overflow-hidden">
      <Table>
        <TableHeader>
          <TableRow className="bg-muted/50">
            <TableHead className="font-semibold">Supplier Name</TableHead>
            <TableHead className="font-semibold">Country</TableHead>
            <TableHead className="font-semibold">Product</TableHead>
            <TableHead className="font-semibold">EC Regulation</TableHead>
            <TableHead className="font-semibold">Certification</TableHead>
            <TableHead className="font-semibold">Expiry Date</TableHead>
            <TableHead className="font-semibold">Cert Status</TableHead>
            <TableHead className="font-semibold">Days to Expiry</TableHead>
          </TableRow>
        </TableHeader>
        <TableBody>
          {certificates.map((cert) => {
            const daysToExpiry = getDaysToExpiry(cert.expiryDate);
            const certStatus = getCertStatus(daysToExpiry);

            return (
              <TableRow key={cert.id} className="hover:bg-muted/30">
                <TableCell className="font-medium">{cert.supplierName || '-'}</TableCell>
                <TableCell>{cert.country || '-'}</TableCell>
                <TableCell>{cert.product || '-'}</TableCell>
                <TableCell className="text-sm max-w-[200px] truncate" title={cert.ecRegulation}>
                  {cert.ecRegulation || '-'}
                </TableCell>
                <TableCell>{cert.certification || '-'}</TableCell>
                <TableCell className="text-sm">{cert.expiryDate || '-'}</TableCell>
                <TableCell>
                  <Badge
                    variant="outline"
                    className={statusStyles[certStatus]}
                  >
                    {statusLabels[certStatus]}
                  </Badge>
                </TableCell>
                <TableCell className={`font-medium ${
                  daysToExpiry !== null && daysToExpiry < 0 ? 'text-red-600' :
                  daysToExpiry !== null && daysToExpiry < 30 ? 'text-orange-600' :
                  'text-green-600'
                }`}>
                  {daysToExpiry !== null ? daysToExpiry : '-'}
                </TableCell>
              </TableRow>
            );
          })}
        </TableBody>
      </Table>
    </div>
  );
};

export default ReviewTable;
