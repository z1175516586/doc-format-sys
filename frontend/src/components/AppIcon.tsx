import { MorphIcon } from "morphicons/react";
import type { IconNode } from "lucide";

interface AppIconProps {
  icon: IconNode;
  size?: number;
  label?: string;
}

/** 全站图标统一走 morphicons，不直接用其它图标组件。 */
export function AppIcon({ icon, size = 18, label }: AppIconProps) {
  return <MorphIcon icon={icon} size={size} strokeWidth={1.75} label={label} />;
}
