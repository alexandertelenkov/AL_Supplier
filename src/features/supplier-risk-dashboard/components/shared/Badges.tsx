import { Badge } from "@/components/ui/badge";
import type { ContractStatus, IssueSeverity, RiskLevel } from "../../types";
import { getReadableTextColor } from "../../utils";

export function RiskBadge({ risk, tone }: { risk: RiskLevel; tone?: string }) {
  if (tone) {
    return (
      <Badge
        variant="outline"
        className="border"
        style={{ backgroundColor: tone, borderColor: tone, color: getReadableTextColor(tone) }}
      >
        {risk === "High" ? "High risk" : risk}
      </Badge>
    );
  }
  if (risk === "High") return <Badge variant="destructive">High risk</Badge>;
  if (risk === "Non-risk") return <Badge variant="secondary">Non-risk</Badge>;
  return <Badge variant="outline">Unknown</Badge>;
}

export function ContractBadge({ status, tone }: { status: ContractStatus; tone?: string }) {
  if (tone) {
    return (
      <Badge
        variant="outline"
        className="border"
        style={{ backgroundColor: tone, borderColor: tone, color: getReadableTextColor(tone) }}
      >
        {status}
      </Badge>
    );
  }
  if (status === "Signed") return <Badge variant="secondary">Signed</Badge>;
  if (status === "Sent") return <Badge variant="outline">Sent</Badge>;
  if (status === "Not compliant") return <Badge variant="destructive">Not compliant</Badge>;
  if (status === "Review") return <Badge variant="outline">Review</Badge>;
  if (status === "N/A") return <Badge variant="secondary">N/A</Badge>;
  return <Badge variant="outline">Not sent</Badge>;
}

export function SeverityBadge({ sev }: { sev: IssueSeverity }) {
  if (sev === "Critical") return <Badge variant="destructive">Critical</Badge>;
  if (sev === "High") return <Badge variant="destructive">High</Badge>;
  if (sev === "Medium") return <Badge variant="outline">Medium</Badge>;
  return <Badge variant="secondary">Low</Badge>;
}
