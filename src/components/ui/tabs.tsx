import React, { createContext, useContext } from 'react'
import { cn } from '@/lib/cn'

type TabsCtx = { value: string; onValueChange: (v: string) => void }
const Ctx = createContext<TabsCtx | null>(null)

export function Tabs({ value, onValueChange, className, children }: { value: string; onValueChange: (v: string) => void; className?: string; children: React.ReactNode }) {
  return (
    <Ctx.Provider value={{ value, onValueChange }}>
      <div className={cn('w-full', className)}>{children}</div>
    </Ctx.Provider>
  )
}

export function TabsList({ className, ...props }: React.HTMLAttributes<HTMLDivElement>) {
  return <div className={cn('rounded-2xl bg-muted p-1', className)} {...props} />
}

export function TabsTrigger({ value, className, children }: { value: string; className?: string; children: React.ReactNode }) {
  const ctx = useContext(Ctx)
  if (!ctx) throw new Error('TabsTrigger must be used inside Tabs')
  const active = ctx.value === value
  return (
    <button
      type="button"
      onClick={() => ctx.onValueChange(value)}
      className={cn(
        'inline-flex items-center justify-center rounded-2xl px-3 py-2 text-sm font-medium transition hover:opacity-90',
        active ? 'bg-background shadow-sm' : 'text-muted-foreground',
        className,
      )}
    >
      {children}
    </button>
  )
}

export function TabsContent({ value, className, children }: { value: string; className?: string; children: React.ReactNode }) {
  const ctx = useContext(Ctx)
  if (!ctx) throw new Error('TabsContent must be used inside Tabs')
  if (ctx.value !== value) return null
  return <div className={cn('mt-2', className)}>{children}</div>
}
