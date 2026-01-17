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
import { FileText, CheckCircle2, AlertTriangle, XCircle, HelpCircle } from 'lucide-react';

interface ReviewTableProps {
  certificates: CertificateData[];
}

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

function getCertStatus(daysToExpiry: number | null): 'expired' | 'expiring_soon' | 'up_to_date' | 'unknown' {
  if (daysToExpiry === null) return 'unknown';
  if (daysToExpiry < 0) return 'expired';
  if (daysToExpiry < 30) return 'expiring_soon';
  return 'up_to_date';
}

const statusConfig = {
  expired: {
    bg: 'bg-red-50',
    text: 'text-red-700',
    border: 'border-red-200',
    icon: XCircle,
    label: 'Expired',
  },
  expiring_soon: {
    bg: 'bg-amber-50',
    text: 'text-amber-700',
    border: 'border-amber-200',
    icon: AlertTriangle,
    label: 'Expiring Soon',
  },
  up_to_date: {
    bg: 'bg-emerald-50',
    text: 'text-emerald-700',
    border: 'border-emerald-200',
    icon: CheckCircle2,
    label: 'Valid',
  },
  unknown: {
    bg: 'bg-gray-50',
    text: 'text-gray-500',
    border: 'border-gray-200',
    icon: HelpCircle,
    label: 'Unknown',
  },
};

const ReviewTable = ({ certificates }: ReviewTableProps) => {
  if (certificates.length === 0) {
    return (
      <div className="border border-gray-200 rounded-xl p-8 sm:p-16 text-center bg-gray-50/50">
        <div className="w-12 h-12 sm:w-16 sm:h-16 rounded-full bg-gray-100 flex items-center justify-center mx-auto mb-3 sm:mb-4">
          <FileText className="w-6 h-6 sm:w-8 sm:h-8 text-gray-400" />
        </div>
        <h4 className="text-sm sm:text-base font-medium text-gray-900 mb-1 sm:mb-2">
          No certificates analyzed yet
        </h4>
        <p className="text-xs sm:text-sm text-gray-500 max-w-sm mx-auto">
          Upload PDF certificates using the drop zone above to begin extracting compliance data.
        </p>
      </div>
    );
  }

  return (
    <div className="border border-gray-200 rounded-xl overflow-hidden min-w-[800px]">
      <Table>
        <TableHeader>
          <TableRow className="bg-gray-50/80 hover:bg-gray-50/80 border-b border-gray-200">
            <TableHead className="font-semibold text-gray-700 text-[10px] sm:text-xs uppercase tracking-wider py-3 sm:py-4 px-2 sm:px-4">
              Supplier
            </TableHead>
            <TableHead className="font-semibold text-gray-700 text-[10px] sm:text-xs uppercase tracking-wider py-3 sm:py-4 px-2 sm:px-4">
              Country
            </TableHead>
            <TableHead className="font-semibold text-gray-700 text-[10px] sm:text-xs uppercase tracking-wider py-3 sm:py-4 px-2 sm:px-4">
              Product
            </TableHead>
            <TableHead className="font-semibold text-gray-700 text-[10px] sm:text-xs uppercase tracking-wider py-3 sm:py-4 px-2 sm:px-4">
              EC Regulation
            </TableHead>
            <TableHead className="font-semibold text-gray-700 text-[10px] sm:text-xs uppercase tracking-wider py-3 sm:py-4 px-2 sm:px-4">
              Cert
            </TableHead>
            <TableHead className="font-semibold text-gray-700 text-[10px] sm:text-xs uppercase tracking-wider py-3 sm:py-4 px-2 sm:px-4">
              Expiry
            </TableHead>
            <TableHead className="font-semibold text-gray-700 text-[10px] sm:text-xs uppercase tracking-wider py-3 sm:py-4 px-2 sm:px-4">
              Status
            </TableHead>
            <TableHead className="font-semibold text-gray-700 text-[10px] sm:text-xs uppercase tracking-wider py-3 sm:py-4 px-2 sm:px-4 text-right">
              Days
            </TableHead>
          </TableRow>
        </TableHeader>
        <TableBody>
          {certificates.map((cert, index) => {
            const daysToExpiry = getDaysToExpiry(cert.expiryDate);
            const certStatus = getCertStatus(daysToExpiry);
            const config = statusConfig[certStatus];
            const StatusIcon = config.icon;

            return (
              <TableRow
                key={cert.id}
                className={`
                  hover:bg-gray-50/50 border-b border-gray-100 last:border-0
                  ${index % 2 === 0 ? 'bg-white' : 'bg-gray-50/30'}
                `}
              >
                <TableCell className="font-medium text-gray-900 py-3 sm:py-4 px-2 sm:px-4 text-xs sm:text-sm">
                  {cert.supplierName || <span className="text-gray-400">-</span>}
                </TableCell>
                <TableCell className="text-gray-600 py-3 sm:py-4 px-2 sm:px-4 text-xs sm:text-sm">
                  {cert.country || <span className="text-gray-400">-</span>}
                </TableCell>
                <TableCell className="max-w-[120px] sm:max-w-[200px] py-3 sm:py-4 px-2 sm:px-4">
                  <span
                    className="block truncate text-gray-600 text-xs sm:text-sm"
                    title={cert.product}
                  >
                    {cert.product || <span className="text-gray-400">-</span>}
                  </span>
                </TableCell>
                <TableCell className="max-w-[120px] sm:max-w-[180px] py-3 sm:py-4 px-2 sm:px-4">
                  <span
                    className="block truncate text-gray-600 text-xs sm:text-sm"
                    title={cert.ecRegulation}
                  >
                    {cert.ecRegulation || <span className="text-gray-400">-</span>}
                  </span>
                </TableCell>
                <TableCell className="text-gray-600 py-3 sm:py-4 px-2 sm:px-4 text-xs sm:text-sm">
                  {cert.certification || <span className="text-gray-400">-</span>}
                </TableCell>
                <TableCell className="text-gray-600 py-3 sm:py-4 px-2 sm:px-4 font-mono text-[10px] sm:text-sm whitespace-nowrap">
                  {cert.expiryDate || <span className="text-gray-400">-</span>}
                </TableCell>
                <TableCell className="py-3 sm:py-4 px-2 sm:px-4">
                  <Badge
                    variant="outline"
                    className={`${config.bg} ${config.text} ${config.border} font-medium gap-1 sm:gap-1.5 px-1.5 sm:px-2.5 py-0.5 sm:py-1 text-[10px] sm:text-xs whitespace-nowrap`}
                  >
                    <StatusIcon className="w-3 h-3 sm:w-3.5 sm:h-3.5" />
                    {config.label}
                  </Badge>
                </TableCell>
                <TableCell className="py-3 sm:py-4 px-2 sm:px-4 text-right">
                  <span className={`
                    font-semibold tabular-nums text-xs sm:text-sm
                    ${daysToExpiry !== null && daysToExpiry < 0 ? 'text-red-600' :
                      daysToExpiry !== null && daysToExpiry < 30 ? 'text-amber-600' :
                      daysToExpiry !== null ? 'text-emerald-600' : 'text-gray-400'}
                  `}>
                    {daysToExpiry !== null ? daysToExpiry : '-'}
                  </span>
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
