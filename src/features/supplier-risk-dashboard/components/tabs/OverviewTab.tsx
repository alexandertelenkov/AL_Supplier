import React from "react";

type OverviewTabProps = {
  children: React.ReactNode;
};

const OverviewTab = React.memo(function OverviewTab({ children }: OverviewTabProps) {
  return <>{children}</>;
});

export default OverviewTab;
