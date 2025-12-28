import React from "react";

type SuppliersTabProps = {
  children: React.ReactNode;
};

const SuppliersTab = React.memo(function SuppliersTab({ children }: SuppliersTabProps) {
  return <>{children}</>;
});

export default SuppliersTab;
