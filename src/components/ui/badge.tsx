import React from 'react'
import { cn } from '@/lib/cn'

type Variant = 'default' | 'secondary' | 'outline' | 'destructive'

export function Badge({ className, variant = 'default', ...props }: React.HTMLAttributes<HTMLSpanElement> & { variant?: Variant }) {
  const v = {
    default: 'bg-primary text-primary-foreground',
    secondary: 'bg-secondary text-secondary-foreground',
    outline: 'border border-input text-foreground',
    destructive: 'bg-destructive text-destructive-foreground',
  }[variant]

  return (
    <span
      className={cn('inline-flex items-center rounded-xl px-2 py-0.5 text-xs font-medium', v, className)}
      {...props}
    />
  )
}
