import React from "react";

type BulkTabProps = {
  children: React.ReactNode;
};

const BulkTab = React.memo(function BulkTab({ children }: BulkTabProps) {
  return <>{children}</>;
});

export default BulkTab;
