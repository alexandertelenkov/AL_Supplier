import React, { createContext, useContext, useEffect, useState } from 'react'
import { cn } from '@/lib/cn'

type DialogCtx = { open: boolean; setOpen: (v: boolean) => void }
const Ctx = createContext<DialogCtx | null>(null)

export function Dialog({ children }: { children: React.ReactNode }) {
  const [open, setOpen] = useState(false)

  // ESC to close
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setOpen(false)
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [])

  return <Ctx.Provider value={{ open, setOpen }}>{children}</Ctx.Provider>
}

export function DialogTrigger({ asChild, children }: { asChild?: boolean; children: React.ReactElement }) {
  const ctx = useContext(Ctx)
  if (!ctx) throw new Error('DialogTrigger must be used inside Dialog')
  const onClick = () => ctx.setOpen(true)
  if (asChild) {
    return React.cloneElement(children, {
      onClick: (e: any) => {
        children.props.onClick?.(e)
        onClick()
      },
    })
  }
  return (
    <button type="button" onClick={onClick}>
      {children}
    </button>
  )
}

export function DialogContent({ className, children }: { className?: string; children: React.ReactNode }) {
  const ctx = useContext(Ctx)
  if (!ctx) throw new Error('DialogContent must be used inside Dialog')
  if (!ctx.open) return null

  return (
    <div className="fixed inset-0 z-50">
      <div
        className="absolute inset-0 bg-black/40"
        onClick={() => ctx.setOpen(false)}
        aria-hidden
      />
      <div className="absolute inset-0 flex items-center justify-center p-4">
        <div className={cn('relative w-full rounded-3xl border bg-background shadow-lg max-h-[90vh] overflow-auto', className)}>
          <button
            type="button"
            className="absolute right-3 top-3 rounded-xl border px-2 py-1 text-xs hover:bg-muted"
            onClick={() => ctx.setOpen(false)}
          >
            Close
          </button>
          <div className="p-4">{children}</div>
        </div>
      </div>
    </div>
  )
}

export function DialogHeader({ className, ...props }: React.HTMLAttributes<HTMLDivElement>) {
  return <div className={cn('mb-3 space-y-1', className)} {...props} />
}

export function DialogTitle({ className, ...props }: React.HTMLAttributes<HTMLHeadingElement>) {
  return <h2 className={cn('text-base font-semibold', className)} {...props} />
}
