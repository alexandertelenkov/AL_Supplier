import React, { createContext, useContext, useMemo } from 'react'
import { cn } from '@/lib/cn'

type Item = { value: string; label: string }

type SelectCtx = {
  value: string
  onValueChange: (v: string) => void
  items: Item[]
}

const Ctx = createContext<SelectCtx | null>(null)

function collectItems(node: React.ReactNode, out: Item[]) {
  React.Children.forEach(node, (child: any) => {
    if (!child) return
    if (Array.isArray(child)) {
      collectItems(child, out)
      return
    }
    if (child?.type && child.type.displayName === 'SelectItem') {
      const value = String(child.props.value)
      const label = typeof child.props.children === 'string' ? child.props.children : String(child.props.children)
      out.push({ value, label })
      return
    }
    if (child?.props?.children) collectItems(child.props.children, out)
  })
}

export function Select({ value, onValueChange, children }: { value: string; onValueChange: (v: string) => void; children: React.ReactNode }) {
  const items = useMemo(() => {
    const out: Item[] = []
    collectItems(children, out)
    // de-dup by value
    const map = new Map<string, Item>()
    for (const it of out) if (!map.has(it.value)) map.set(it.value, it)
    return Array.from(map.values())
  }, [children])

  return <Ctx.Provider value={{ value: String(value), onValueChange, items }}>{children}</Ctx.Provider>
}

export function SelectTrigger({ className, children, ...props }: React.SelectHTMLAttributes<HTMLSelectElement> & { children?: React.ReactNode }) {
  const ctx = useContext(Ctx)
  if (!ctx) throw new Error('SelectTrigger must be used inside Select')
  return (
    <select
      className={cn(
        'flex h-10 w-full rounded-2xl border border-input bg-background px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-ring',
        className,
      )}
      value={ctx.value}
      onChange={(e) => ctx.onValueChange(e.target.value)}
      {...props}
    >
      {ctx.items.map((it) => (
        <option key={it.value} value={it.value}>
          {it.label}
        </option>
      ))}
      {children /* ignored in this lightweight implementation */}
    </select>
  )
}

export function SelectValue(_props: { placeholder?: string }) {
  return null
}

export function SelectContent(_props: { children?: React.ReactNode }) {
  return null
}

export function SelectItem(_props: { value: string; children: React.ReactNode }) {
  return null
}
SelectItem.displayName = 'SelectItem'
