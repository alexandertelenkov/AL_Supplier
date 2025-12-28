import React from "react";

type SettingsTabProps = {
  children: React.ReactNode;
};

const SettingsTab = React.memo(function SettingsTab({ children }: SettingsTabProps) {
  return <>{children}</>;
});

export default SettingsTab;
