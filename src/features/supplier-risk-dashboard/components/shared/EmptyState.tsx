import { AlertCircle } from "lucide-react";
import { Card, CardContent } from "@/components/ui/card";

export function EmptyState({ title, subtitle }: { title: string; subtitle?: string }) {
  return (
    <Card>
      <CardContent className="py-10 text-center">
        <div className="mx-auto mb-3 w-fit rounded-full bg-muted p-3">
          <AlertCircle className="h-5 w-5" />
        </div>
        <div className="text-lg font-semibold">{title}</div>
        {subtitle ? <div className="mt-1 text-sm text-muted-foreground">{subtitle}</div> : null}
      </CardContent>
    </Card>
  );
}
