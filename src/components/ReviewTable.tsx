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

const statusStyles = {
  valid: 'bg-emerald-500/10 text-emerald-600 border-emerald-500/20',
  expired: 'bg-destructive/10 text-destructive border-destructive/20',
  pending: 'bg-amber-500/10 text-amber-600 border-amber-500/20',
  unknown: 'bg-muted text-muted-foreground border-border',
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
            <TableHead className="font-semibold">Certification</TableHead>
            <TableHead className="font-semibold">Dates</TableHead>
            <TableHead className="font-semibold">Status</TableHead>
          </TableRow>
        </TableHeader>
        <TableBody>
          {certificates.map((cert) => (
            <TableRow key={cert.id} className="hover:bg-muted/30">
              <TableCell className="font-medium">{cert.supplierName || '-'}</TableCell>
              <TableCell>{cert.country || '-'}</TableCell>
              <TableCell>{cert.product || '-'}</TableCell>
              <TableCell>{cert.certification || '-'}</TableCell>
              <TableCell className="text-sm">
                <div>{cert.issueDate || '-'} - {cert.expiryDate || '-'}</div>
              </TableCell>
              <TableCell>
                <Badge
                  variant="outline"
                  className={statusStyles[cert.status]}
                >
                  {cert.status.charAt(0).toUpperCase() + cert.status.slice(1)}
                </Badge>
              </TableCell>
            </TableRow>
          ))}
        </TableBody>
      </Table>
    </div>
  );
};

export default ReviewTable;
