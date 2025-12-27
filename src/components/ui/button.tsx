import React from 'react'
import { cn } from '@/lib/cn'

type Variant = 'default' | 'secondary' | 'outline' | 'destructive'
type Size = 'default' | 'sm'

export type ButtonProps = React.ButtonHTMLAttributes<HTMLButtonElement> & {
  variant?: Variant
  size?: Size
}

const variantCls: Record<Variant, string> = {
  default: 'bg-primary text-primary-foreground hover:opacity-90',
  secondary: 'bg-secondary text-secondary-foreground hover:opacity-90',
  outline: 'border border-input bg-background hover:bg-muted',
  destructive: 'bg-destructive text-destructive-foreground hover:opacity-90',
}

const sizeCls: Record<Size, string> = {
  default: 'h-10 px-4 py-2 text-sm',
  sm: 'h-9 px-3 text-sm',
}

export function Button({ className, variant = 'default', size = 'default', ...props }: ButtonProps) {
  return (
    <button
      className={cn(
        'inline-flex items-center justify-center whitespace-nowrap rounded-2xl font-medium transition focus:outline-none focus:ring-2 focus:ring-ring disabled:pointer-events-none disabled:opacity-50',
        variantCls[variant],
        sizeCls[size],
        className,
      )}
      {...props}
    />
  )
}
