"use client"

import { useState, useMemo } from "react"
import { Button } from "@/components/ui/button"
import { Input } from "@/components/ui/input"
import { Label } from "@/components/ui/label"
import { Textarea } from "@/components/ui/textarea"
import { Card } from "@/components/ui/card"
import { FileText, Download, History, User, Plus, Trash2, ChevronLeft, ChevronRight } from "lucide-react"

interface LineItem {
  id: string
  productName: string
  drawingNumber: string
  quantity: string
  unitCost: string
  customsDuty: string
  material: string
  unitWeight: string
}

export default function CommercialProposalPage() {
  const [sidebarCollapsed, setSidebarCollapsed] = useState(false)

  const [globalData, setGlobalData] = useState({
    companyName: "",
    tenderNumber: "",
    deliveryAddress: "",
    deliveryDays: "",
    logisticsCost: "",
    comments: "",
  })

  const [lineItems, setLineItems] = useState<LineItem[]>([
    {
      id: "1",
      productName: "",
      drawingNumber: "",
      quantity: "",
      unitCost: "",
      customsDuty: "",
      material: "",
      unitWeight: "",
    },
  ])

  const [margin, setMargin] = useState(0)

  const totalWeight = useMemo(() => {
    return lineItems.reduce((sum, item) => {
      const quantity = Number.parseFloat(item.quantity) || 0
      const unitWeight = Number.parseFloat(item.unitWeight) || 0
      return sum + quantity * unitWeight
    }, 0)
  }, [lineItems])

  const handleGlobalChange = (field: string, value: string) => {
    setGlobalData((prev) => ({ ...prev, [field]: value }))
  }

  const handleLineItemChange = (id: string, field: keyof LineItem, value: string) => {
    setLineItems((prev) => prev.map((item) => (item.id === id ? { ...item, [field]: value } : item)))
  }

  const addLineItem = () => {
    const newItem: LineItem = {
      id: Date.now().toString(),
      productName: "",
      drawingNumber: "",
      quantity: "",
      unitCost: "",
      customsDuty: "",
      material: "",
      unitWeight: "",
    }
    setLineItems((prev) => [...prev, newItem])
  }

  const removeLineItem = (id: string) => {
    if (lineItems.length > 1) {
      setLineItems((prev) => prev.filter((item) => item.id !== id))
    }
  }

  return (
    <div className="flex min-h-screen bg-background">
      {/* Sidebar */}
      <aside
        className={`relative border-r border-border bg-card transition-all duration-300 ${
          sidebarCollapsed ? "w-0 overflow-hidden" : "w-56"
        }`}
      >
        <div className="flex h-16 items-center border-b border-border px-6">
          <h1 className="font-sans text-lg font-semibold tracking-tight text-foreground">КП Система</h1>
        </div>
        <nav className="space-y-1 p-4">
          <button className="flex w-full items-center gap-3 rounded-lg bg-accent px-4 py-2.5 text-sm font-medium text-accent-foreground transition-colors">
            <FileText className="h-4 w-4" />
            Логистика
          </button>
          <button className="flex w-full items-center gap-3 rounded-lg px-4 py-2.5 text-sm font-medium text-muted-foreground transition-colors hover:bg-secondary hover:text-foreground">
            <FileText className="h-4 w-4" />
            Пошлина
          </button>
          <button className="flex w-full items-center gap-3 rounded-lg px-4 py-2.5 text-sm font-medium text-muted-foreground transition-colors hover:bg-secondary hover:text-foreground">
            <FileText className="h-4 w-4" />
            Аналоги по стандарту GB
          </button>
          <button className="flex w-full items-center gap-3 rounded-lg px-4 py-2.5 text-sm font-medium text-muted-foreground transition-colors hover:bg-secondary hover:text-foreground">
            <FileText className="h-4 w-4" />
            Цена металла
          </button>
          <button className="flex w-full items-center gap-3 rounded-lg px-4 py-2.5 text-sm font-medium text-muted-foreground transition-colors hover:bg-secondary hover:text-foreground">
            <FileText className="h-4 w-4" />
            Крановые колеса
          </button>
        </nav>
      </aside>

      <button
        onClick={() => setSidebarCollapsed(!sidebarCollapsed)}
        className="absolute left-0 top-20 z-10 flex h-10 w-6 items-center justify-center rounded-r-lg border border-l-0 border-border bg-card text-muted-foreground transition-all hover:bg-accent hover:text-accent-foreground"
        style={{ left: sidebarCollapsed ? "0" : "224px" }}
        aria-label={sidebarCollapsed ? "Развернуть панель" : "Свернуть панель"}
      >
        {sidebarCollapsed ? <ChevronRight className="h-4 w-4" /> : <ChevronLeft className="h-4 w-4" />}
      </button>

      {/* Main Content */}
      <main className="flex-1">
        {/* Header */}
        <header className="flex h-16 items-center justify-between border-b border-border bg-card px-8">
          <div className="flex items-center gap-8">
            <h2 className="font-sans text-xl font-semibold tracking-tight text-foreground">
              Создание коммерческого предложения
            </h2>
          </div>
          <div className="flex items-center gap-4">
            <button className="flex items-center gap-2 text-sm text-muted-foreground transition-colors hover:text-foreground">
              <History className="h-4 w-4" />
              История
            </button>
            <button className="flex items-center gap-2 text-sm text-muted-foreground transition-colors hover:text-foreground">
              <User className="h-4 w-4" />
              tsc_perm@mail.ru
            </button>
          </div>
        </header>

        {/* Form Content */}
        <div className="mx-auto max-w-7xl p-8">
          <div className="space-y-8">
            <section>
              <h3 className="mb-6 font-sans text-base font-semibold text-foreground">Общая информация</h3>
              <div className="grid gap-6 md:grid-cols-3">
                <div className="space-y-2">
                  <Label htmlFor="companyName" className="text-sm font-medium">
                    Название компании
                  </Label>
                  <Input
                    id="companyName"
                    value={globalData.companyName}
                    onChange={(e) => handleGlobalChange("companyName", e.target.value)}
                    className="h-11"
                  />
                </div>
                <div className="space-y-2">
                  <Label htmlFor="tenderNumber" className="text-sm font-medium">
                    Номер тендера
                  </Label>
                  <Input
                    id="tenderNumber"
                    value={globalData.tenderNumber}
                    onChange={(e) => handleGlobalChange("tenderNumber", e.target.value)}
                    className="h-11"
                  />
                </div>
                <div className="space-y-2">
                  <Label className="text-sm font-medium">Итоговый вес</Label>
                  <div className="flex h-11 items-center rounded-lg border border-border bg-muted/30 px-4">
                    <span className="font-mono text-sm font-semibold text-foreground">{totalWeight.toFixed(2)} кг</span>
                  </div>
                </div>
                <div className="space-y-2 md:col-span-3">
                  <Label htmlFor="deliveryAddress" className="text-sm font-medium">
                    Адрес доставки
                  </Label>
                  <Input
                    id="deliveryAddress"
                    value={globalData.deliveryAddress}
                    onChange={(e) => handleGlobalChange("deliveryAddress", e.target.value)}
                    className="h-11"
                  />
                </div>
                <div className="space-y-2">
                  <Label htmlFor="deliveryDays" className="text-sm font-medium">
                    Срок поставки (календарных дней)
                  </Label>
                  <Input
                    id="deliveryDays"
                    type="number"
                    value={globalData.deliveryDays}
                    onChange={(e) => handleGlobalChange("deliveryDays", e.target.value)}
                    className="h-11"
                  />
                </div>
                <div className="space-y-2">
                  <Label htmlFor="logisticsCost" className="text-sm font-medium">
                    Стоимость логистики (руб.)
                  </Label>
                  <Input
                    id="logisticsCost"
                    type="number"
                    step="0.01"
                    value={globalData.logisticsCost}
                    onChange={(e) => handleGlobalChange("logisticsCost", e.target.value)}
                    className="h-11"
                  />
                </div>
              </div>
            </section>

            <section>
              <div className="mb-6 flex items-center justify-between">
                <h3 className="font-sans text-base font-semibold text-foreground">Позиции товаров</h3>
                <Button onClick={addLineItem} variant="outline" size="sm" className="gap-2 bg-transparent">
                  <Plus className="h-4 w-4" />
                  Добавить позицию
                </Button>
              </div>

              <div className="space-y-6">
                {lineItems.map((item, index) => (
                  <Card key={item.id} className="relative p-6">
                    {lineItems.length > 1 && (
                      <button
                        onClick={() => removeLineItem(item.id)}
                        className="absolute right-4 top-4 rounded-lg p-2 text-muted-foreground transition-colors hover:bg-destructive/10 hover:text-destructive"
                        aria-label="Удалить позицию"
                      >
                        <Trash2 className="h-4 w-4" />
                      </button>
                    )}

                    <div className="mb-4">
                      <span className="text-sm font-medium text-muted-foreground">Позиция {index + 1}</span>
                    </div>

                    <div className="grid gap-6 md:grid-cols-3 lg:grid-cols-4">
                      <div className="space-y-2">
                        <Label htmlFor={`productName-${item.id}`} className="text-sm font-medium">
                          Наименование товара
                        </Label>
                        <Input
                          id={`productName-${item.id}`}
                          value={item.productName}
                          onChange={(e) => handleLineItemChange(item.id, "productName", e.target.value)}
                          className="h-11"
                        />
                      </div>
                      <div className="space-y-2">
                        <Label htmlFor={`drawingNumber-${item.id}`} className="text-sm font-medium">
                          Номер чертежа
                        </Label>
                        <Input
                          id={`drawingNumber-${item.id}`}
                          value={item.drawingNumber}
                          onChange={(e) => handleLineItemChange(item.id, "drawingNumber", e.target.value)}
                          className="h-11"
                        />
                      </div>
                      <div className="space-y-2">
                        <Label htmlFor={`quantity-${item.id}`} className="text-sm font-medium">
                          Количество
                        </Label>
                        <Input
                          id={`quantity-${item.id}`}
                          type="number"
                          value={item.quantity}
                          onChange={(e) => handleLineItemChange(item.id, "quantity", e.target.value)}
                          className="h-11"
                        />
                      </div>
                      <div className="space-y-2">
                        <Label htmlFor={`unitCost-${item.id}`} className="text-sm font-medium">
                          Закупочная стоимость за шт. (юани)
                        </Label>
                        <Input
                          id={`unitCost-${item.id}`}
                          type="number"
                          step="0.01"
                          value={item.unitCost}
                          onChange={(e) => handleLineItemChange(item.id, "unitCost", e.target.value)}
                          className="h-11"
                        />
                      </div>
                      <div className="space-y-2">
                        <Label htmlFor={`customsDuty-${item.id}`} className="text-sm font-medium">
                          Процент пошлины (%)
                        </Label>
                        <Input
                          id={`customsDuty-${item.id}`}
                          type="number"
                          step="0.1"
                          value={item.customsDuty}
                          onChange={(e) => handleLineItemChange(item.id, "customsDuty", e.target.value)}
                          className="h-11"
                        />
                      </div>
                      <div className="space-y-2">
                        <Label htmlFor={`material-${item.id}`} className="text-sm font-medium">
                          Материал
                        </Label>
                        <Input
                          id={`material-${item.id}`}
                          value={item.material}
                          onChange={(e) => handleLineItemChange(item.id, "material", e.target.value)}
                          className="h-11"
                        />
                      </div>
                      <div className="space-y-2">
                        <Label htmlFor={`unitWeight-${item.id}`} className="text-sm font-medium">
                          Вес за шт. (кг)
                        </Label>
                        <Input
                          id={`unitWeight-${item.id}`}
                          type="number"
                          step="0.1"
                          value={item.unitWeight}
                          onChange={(e) => handleLineItemChange(item.id, "unitWeight", e.target.value)}
                          className="h-11"
                        />
                      </div>
                    </div>
                  </Card>
                ))}
              </div>
            </section>

            {/* Comments Section */}
            <section>
              <div className="space-y-2">
                <Label htmlFor="comments" className="text-sm font-medium">
                  Комментарий
                </Label>
                <Textarea
                  id="comments"
                  placeholder="Дополнительная информация для истории..."
                  value={globalData.comments}
                  onChange={(e) => handleGlobalChange("comments", e.target.value)}
                  className="min-h-[100px] resize-none"
                />
              </div>
            </section>

            {/* Margin Display */}
            <Card className="border-2 bg-muted/30 p-8">
              <div className="flex items-center justify-between">
                <div>
                  <p className="mb-1 text-sm font-medium text-muted-foreground">Итоговая маржа</p>
                  <p className="text-xs text-muted-foreground">Желаемый процент прибыли</p>
                </div>
                <div className="flex items-center gap-4">
                  <div className="flex h-16 w-32 items-center justify-center rounded-lg border-2 border-border bg-card">
                    <span className="font-mono text-3xl font-bold text-foreground">{margin.toFixed(1)}</span>
                  </div>
                  <div className="flex h-16 w-16 items-center justify-center rounded-lg bg-primary">
                    <span className="font-mono text-2xl font-bold text-primary-foreground">%</span>
                  </div>
                </div>
              </div>
            </Card>

            {/* Action Buttons */}
            <div className="flex items-center justify-between border-t border-border pt-8">
              <p className="text-sm text-muted-foreground">Будет скачан ZIP-архив, содержащий файлы Excel и Word.</p>
              <Button size="lg" className="h-12 gap-2 px-8 font-medium">
                <Download className="h-4 w-4" />
                Сгенерировать КП
              </Button>
            </div>
          </div>
        </div>
      </main>
    </div>
  )
}
