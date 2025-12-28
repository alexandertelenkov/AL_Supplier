import React from "react";

type DuplicatesTabProps = {
  children: React.ReactNode;
};

const DuplicatesTab = React.memo(function DuplicatesTab({ children }: DuplicatesTabProps) {
  return <>{children}</>;
});

export default DuplicatesTab;
