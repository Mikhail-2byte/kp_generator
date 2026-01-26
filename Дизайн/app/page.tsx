"use client"

import React from "react"

import { useState } from "react"
import { Button } from "@/components/ui/button"
import { Badge } from "@/components/ui/badge"
import { VariantStepper } from "@/components/proposal-form/variant-stepper"
import { VariantSidebar } from "@/components/proposal-form/variant-sidebar"
import { VariantAccordion } from "@/components/proposal-form/variant-accordion"
import { 
  Layers,
  PanelRightClose,
  LayoutList,
  ChevronLeft,
  ChevronRight
} from "lucide-react"

type VariantType = "stepper" | "sidebar" | "accordion"

const variants: { id: VariantType; name: string; description: string; icon: React.ReactNode }[] = [
  { 
    id: "stepper", 
    name: "Пошаговый мастер", 
    description: "Последовательное заполнение с навигацией по шагам",
    icon: <Layers className="w-5 h-5" />
  },
  { 
    id: "sidebar", 
    name: "Боковая панель", 
    description: "Форма слева, расчёт справа — всё на одном экране",
    icon: <PanelRightClose className="w-5 h-5" />
  },
  { 
    id: "accordion", 
    name: "Аккордеон", 
    description: "Сворачиваемые секции с прикреплённой карточкой итогов",
    icon: <LayoutList className="w-5 h-5" />
  },
]

export default function Page() {
  const [currentVariant, setCurrentVariant] = useState<VariantType>("stepper")
  const [showSelector, setShowSelector] = useState(true)

  const currentIndex = variants.findIndex(v => v.id === currentVariant)
  
  const goToPrevious = () => {
    const newIndex = currentIndex === 0 ? variants.length - 1 : currentIndex - 1
    setCurrentVariant(variants[newIndex].id)
  }
  
  const goToNext = () => {
    const newIndex = currentIndex === variants.length - 1 ? 0 : currentIndex + 1
    setCurrentVariant(variants[newIndex].id)
  }

  return (
    <div className="min-h-screen">
      {/* Floating Variant Selector */}
      {showSelector && (
        <div className="fixed bottom-6 left-1/2 -translate-x-1/2 z-50">
          <div className="bg-card border shadow-2xl rounded-2xl p-2 flex items-center gap-2">
            <Button 
              variant="ghost" 
              size="icon"
              onClick={goToPrevious}
              className="shrink-0"
            >
              <ChevronLeft className="w-4 h-4" />
            </Button>
            
            <div className="flex items-center gap-2 px-2">
              {variants.map((variant) => (
                <button
                  key={variant.id}
                  onClick={() => setCurrentVariant(variant.id)}
                  className={`flex items-center gap-3 px-4 py-3 rounded-xl transition-all ${
                    currentVariant === variant.id 
                      ? "bg-primary text-primary-foreground" 
                      : "hover:bg-muted"
                  }`}
                >
                  {variant.icon}
                  <div className="text-left">
                    <p className="font-medium text-sm whitespace-nowrap">{variant.name}</p>
                    <p className={`text-xs ${
                      currentVariant === variant.id 
                        ? "text-primary-foreground/70" 
                        : "text-muted-foreground"
                    }`}>
                      Вариант {variants.findIndex(v => v.id === variant.id) + 1}
                    </p>
                  </div>
                </button>
              ))}
            </div>

            <Button 
              variant="ghost" 
              size="icon"
              onClick={goToNext}
              className="shrink-0"
            >
              <ChevronRight className="w-4 h-4" />
            </Button>
          </div>
        </div>
      )}

      {/* Toggle Button */}
      <button
        onClick={() => setShowSelector(!showSelector)}
        className="fixed bottom-6 right-6 z-50 bg-card border shadow-lg rounded-full px-4 py-2 text-sm font-medium hover:bg-muted transition-colors flex items-center gap-2"
      >
        <Badge variant="secondary" className="text-xs">
          {currentIndex + 1}/{variants.length}
        </Badge>
        {showSelector ? "Скрыть выбор" : "Показать варианты"}
      </button>

      {/* Render Current Variant */}
      {currentVariant === "stepper" && <VariantStepper />}
      {currentVariant === "sidebar" && <VariantSidebar />}
      {currentVariant === "accordion" && <VariantAccordion />}
    </div>
  )
}
