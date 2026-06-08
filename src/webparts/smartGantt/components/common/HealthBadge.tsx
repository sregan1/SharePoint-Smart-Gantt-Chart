import * as React from 'react';
import { TaskHealth, ProjectHealth } from '../../models';
import { healthColor, healthLightColor, healthLabel } from '../../utils/healthUtils';

interface IHealthBadgeProps {
  health: TaskHealth | ProjectHealth;
  size?: 'sm' | 'md';
}

export const HealthBadge: React.FC<IHealthBadgeProps> = ({ health, size = 'sm' }) => {
  const color = healthColor(health);
  const bg = healthLightColor(health);
  const label = healthLabel(health);

  if (size === 'sm') {
    return (
      <span style={{
        display: 'inline-flex',
        alignItems: 'center',
        gap: 4,
        fontSize: 11,
        fontWeight: 600,
        color,
        whiteSpace: 'nowrap',
      }}>
        <span style={{
          width: 6,
          height: 6,
          borderRadius: '50%',
          backgroundColor: color,
          flexShrink: 0,
          display: 'inline-block',
        }} />
        {label}
      </span>
    );
  }

  return (
    <span style={{
      display: 'inline-flex',
      alignItems: 'center',
      gap: 5,
      padding: '3px 8px',
      borderRadius: 12,
      backgroundColor: bg,
      border: `1px solid ${color}22`,
      fontSize: 12,
      fontWeight: 600,
      color,
      whiteSpace: 'nowrap',
    }}>
      <span style={{
        width: 7,
        height: 7,
        borderRadius: '50%',
        backgroundColor: color,
        flexShrink: 0,
        display: 'inline-block',
      }} />
      {label}
    </span>
  );
};
