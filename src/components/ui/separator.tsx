import React from 'react'
import { cn } from '@/lib/cn'

export function Separator({ className, ...props }: React.HTMLAttributes<HTMLHRElement>) {
  return <hr className={cn('my-2 border-t border-border', className)} {...props} />
}
