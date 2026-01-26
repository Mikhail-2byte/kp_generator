"use client"

import { useState } from "react"
import { Button } from "@/components/ui/button"
import { Input } from "@/components/ui/input"
import { Label } from "@/components/ui/label"
import { Textarea } from "@/components/ui/textarea"
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select"
import { Badge } from "@/components/ui/badge"
import { Separator } from "@/components/ui/separator"
import { ScrollArea } from "@/components/ui/scroll-area"
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs"
import { Switch } from "@/components/ui/switch"
import { 
  FileText,
  Building2, 
  Truck, 
  Package,
  Plus,
  Trash2,
  Upload,
  RotateCcw,
  Save,
  Download,
  TrendingUp,
  DollarSign,
  Percent,
  Scale
} from "lucide-react"

interface Product {
  id: number
  name: string
  drawing: string
  material: string
  purchasePrice: number
  quantity: number
  weight: number
  duty: number
  margin: number
}

export function VariantSidebar() {
  const [useCommonMargin, setUseCommonMargin] = useState(true)
  const [products, setProducts] = useState<Product[]>([
    { 
      id: 1, 
      name: "Деталь литая", 
      drawing: "ЧРТ-001",
      material: "Сталь 45", 
      purchasePrice: 24480,
      quantity: 10, 
      weight: 102,
      duty: 5,
      margin: 25
    }
  ])

  const addProduct = () => {
    setProducts([...products, { 
      id: Date.now(), 
      name: "", 
      drawing: "",
      material: "", 
      purchasePrice: 0,
      quantity: 1, 
      weight: 0,
      duty: 5,
      margin: 25
    }])
  }

  const removeProduct = (id: number) => {
    if (products.length > 1) {
      setProducts(products.filter(p => p.id !== id))
    }
  }

  return (
    <div className="min-h-screen bg-background flex">
      {/* Left Panel - Form */}
      <div className="flex-1 flex flex-col">
        {/* Top Bar */}
        <header className="h-16 border-b bg-card flex items-center justify-between px-6">
          <div className="flex items-center gap-3">
            <div className="w-8 h-8 rounded-lg bg-primary flex items-center justify-center">
              <FileText className="w-4 h-4 text-primary-foreground" />
            </div>
            <div>
              <h1 className="font-semibold text-foreground">Коммерческое предложение</h1>
              <p className="text-xs text-muted-foreground">Тендер ИП1434</p>
            </div>
          </div>
          <div className="flex items-center gap-2">
            <Button variant="ghost" size="sm">
              <RotateCcw className="w-4 h-4 mr-2" />
              Сбросить
            </Button>
            <Button variant="outline" size="sm">
              <Save className="w-4 h-4 mr-2" />
              Сохранить
            </Button>
            <Button size="sm">
              <Download className="w-4 h-4 mr-2" />
              Сгенерировать КП
            </Button>
          </div>
        </header>

        {/* Main Content */}
        <ScrollArea className="flex-1">
          <div className="p-6 space-y-8">
            {/* Company & Delivery Section */}
            <section className="space-y-4">
              <div className="flex items-center gap-2">
                <Building2 className="w-5 h-5 text-primary" />
                <h2 className="text-lg font-semibold">Клиент и условия</h2>
              </div>
              
              <div className="grid grid-cols-3 gap-4">
                <div className="space-y-2">
                  <Label className="text-xs text-muted-foreground uppercase tracking-wider">Регион бюджета</Label>
                  <Select defaultValue="china">
                    <SelectTrigger>
                      <SelectValue />
                    </SelectTrigger>
                    <SelectContent>
                      <SelectItem value="china">Юани (Китай)</SelectItem>
                      <SelectItem value="russia">Рубли (РФ)</SelectItem>
                    </SelectContent>
                  </Select>
                </div>
                <div className="space-y-2">
                  <Label className="text-xs text-muted-foreground uppercase tracking-wider">Номер тендера</Label>
                  <Input defaultValue="ИП1434" />
                </div>
                <div className="space-y-2">
                  <Label className="text-xs text-muted-foreground uppercase tracking-wider">Срок действия</Label>
                  <div className="flex items-center gap-2">
                    <Input type="number" defaultValue="30" className="flex-1" />
                    <span className="text-sm text-muted-foreground">дней</span>
                  </div>
                </div>
              </div>

              <div className="grid grid-cols-2 gap-4">
                <div className="space-y-2">
                  <Label className="text-xs text-muted-foreground uppercase tracking-wider">Компания</Label>
                  <Input defaultValue='ОАО "Кузлитмаш"' />
                </div>
                <div className="space-y-2">
                  <Label className="text-xs text-muted-foreground uppercase tracking-wider">Условия оплаты</Label>
                  <Input defaultValue="100% по факту поставки, 30 дней" />
                </div>
              </div>

              <div className="space-y-2">
                <Label className="text-xs text-muted-foreground uppercase tracking-wider">Адрес доставки</Label>
                <Textarea 
                  defaultValue="Республика Беларусь, Брестская обл., г. Пинск, 225710, пр. Жолтовского, 109"
                  rows={2}
                  className="resize-none"
                />
              </div>

              <div className="grid grid-cols-2 gap-4">
                <div className="space-y-2">
                  <Label className="text-xs text-muted-foreground uppercase tracking-wider">Срок поставки</Label>
                  <div className="flex items-center gap-2">
                    <Input type="number" defaultValue="150" className="flex-1" />
                    <span className="text-sm text-muted-foreground">дней</span>
                  </div>
                </div>
                <div className="space-y-2">
                  <Label className="text-xs text-muted-foreground uppercase tracking-wider">Гарантия</Label>
                  <Input defaultValue="12 месяцев с ввода в эксплуатацию" />
                </div>
              </div>
            </section>

            <Separator />

            {/* Logistics Section */}
            <section className="space-y-4">
              <div className="flex items-center gap-2">
                <Truck className="w-5 h-5 text-primary" />
                <h2 className="text-lg font-semibold">Логистика</h2>
              </div>

              <Tabs defaultValue="truck" className="w-full">
                <TabsList className="grid w-full grid-cols-2">
                  <TabsTrigger value="truck">Фура</TabsTrigger>
                  <TabsTrigger value="trailer">Трал</TabsTrigger>
                </TabsList>
              </Tabs>

              <div className="grid grid-cols-4 gap-4">
                <div className="space-y-2">
                  <Label className="text-xs text-muted-foreground uppercase tracking-wider">Вес (кг)</Label>
                  <Input type="number" defaultValue="1020" />
                </div>
                <div className="space-y-2">
                  <Label className="text-xs text-muted-foreground uppercase tracking-wider">Габарит</Label>
                  <Input placeholder="ДхШхВ" />
                </div>
                <div className="space-y-2">
                  <Label className="text-xs text-muted-foreground uppercase tracking-wider">Оценочный вес</Label>
                  <Input type="number" defaultValue="1020" disabled className="bg-muted" />
                </div>
                <div className="space-y-2">
                  <Label className="text-xs text-muted-foreground uppercase tracking-wider">Город</Label>
                  <Select defaultValue="brest">
                    <SelectTrigger>
                      <SelectValue />
                    </SelectTrigger>
                    <SelectContent>
                      <SelectItem value="brest">Брест</SelectItem>
                      <SelectItem value="minsk">Минск</SelectItem>
                      <SelectItem value="gomel">Гомель</SelectItem>
                    </SelectContent>
                  </Select>
                </div>
              </div>

              <div className="flex items-center justify-between p-4 bg-muted/50 rounded-lg">
                <div>
                  <p className="text-sm text-muted-foreground">Стоимость логистики</p>
                  <p className="text-xl font-semibold">150 000 ₽</p>
                </div>
                <Button variant="secondary" size="sm">
                  Пересчитать
                </Button>
              </div>
            </section>

            <Separator />

            {/* Products Section */}
            <section className="space-y-4">
              <div className="flex items-center justify-between">
                <div className="flex items-center gap-2">
                  <Package className="w-5 h-5 text-primary" />
                  <h2 className="text-lg font-semibold">Товары и услуги</h2>
                </div>
                <div className="flex items-center gap-4">
                  <Button variant="outline" size="sm">
                    <Upload className="w-4 h-4 mr-2" />
                    Импорт из Excel
                  </Button>
                </div>
              </div>

              <div className="flex items-center justify-between p-3 bg-secondary rounded-lg">
                <div className="flex items-center gap-2">
                  <Switch 
                    checked={useCommonMargin} 
                    onCheckedChange={setUseCommonMargin}
                  />
                  <Label className="font-medium">Общая маржа для всех позиций</Label>
                </div>
                {useCommonMargin && (
                  <div className="flex items-center gap-2">
                    <Input 
                      type="number" 
                      defaultValue="25" 
                      className="w-20 text-center"
                    />
                    <span className="text-muted-foreground">%</span>
                  </div>
                )}
              </div>

              {/* Products Table Header */}
              <div className="grid grid-cols-12 gap-2 px-4 py-2 bg-muted rounded-t-lg text-xs font-medium text-muted-foreground uppercase tracking-wider">
                <div className="col-span-3">Наименование</div>
                <div className="col-span-2">Материал</div>
                <div>Кол-во</div>
                <div>Вес (кг)</div>
                <div className="col-span-2">Цена закупа</div>
                <div>Пошлина</div>
                <div>Маржа</div>
                <div></div>
              </div>

              {/* Products Rows */}
              {products.map((product) => (
                <div key={product.id} className="grid grid-cols-12 gap-2 px-4 py-3 border rounded-lg items-center">
                  <div className="col-span-3">
                    <Input 
                      defaultValue={product.name} 
                      placeholder="Название"
                      className="h-9"
                    />
                  </div>
                  <div className="col-span-2">
                    <Input 
                      defaultValue={product.material} 
                      placeholder="Материал"
                      className="h-9"
                    />
                  </div>
                  <div>
                    <Input 
                      type="number" 
                      defaultValue={product.quantity}
                      className="h-9 text-center"
                    />
                  </div>
                  <div>
                    <Input 
                      type="number" 
                      defaultValue={product.weight}
                      className="h-9 text-center"
                    />
                  </div>
                  <div className="col-span-2">
                    <Input 
                      type="number" 
                      defaultValue={product.purchasePrice}
                      className="h-9"
                    />
                  </div>
                  <div>
                    <Input 
                      type="number" 
                      defaultValue={product.duty}
                      className="h-9 text-center"
                      disabled={useCommonMargin}
                    />
                  </div>
                  <div>
                    <Input 
                      type="number" 
                      defaultValue={product.margin}
                      className="h-9 text-center"
                      disabled={useCommonMargin}
                    />
                  </div>
                  <div className="flex justify-center">
                    <Button 
                      variant="ghost" 
                      size="icon"
                      className="h-9 w-9 text-muted-foreground hover:text-destructive"
                      onClick={() => removeProduct(product.id)}
                    >
                      <Trash2 className="w-4 h-4" />
                    </Button>
                  </div>
                </div>
              ))}

              <Button variant="outline" className="w-full bg-transparent" onClick={addProduct}>
                <Plus className="w-4 h-4 mr-2" />
                Добавить позицию
              </Button>
            </section>
          </div>
        </ScrollArea>
      </div>

      {/* Right Panel - Summary */}
      <aside className="w-80 border-l bg-card flex flex-col">
        <div className="p-6 border-b">
          <h3 className="font-semibold text-lg">Итоговый расчёт</h3>
          <p className="text-sm text-muted-foreground">Предварительная калькуляция</p>
        </div>

        <ScrollArea className="flex-1">
          <div className="p-6 space-y-6">
            {/* Key Metrics */}
            <div className="grid grid-cols-2 gap-3">
              <div className="p-4 bg-primary/5 rounded-lg border border-primary/10">
                <div className="flex items-center gap-2 text-primary mb-1">
                  <DollarSign className="w-4 h-4" />
                  <span className="text-xs font-medium uppercase">Итого</span>
                </div>
                <p className="text-2xl font-bold">165 300</p>
                <p className="text-xs text-muted-foreground">юаней</p>
              </div>
              <div className="p-4 bg-accent/10 rounded-lg">
                <div className="flex items-center gap-2 text-accent mb-1">
                  <TrendingUp className="w-4 h-4" />
                  <span className="text-xs font-medium uppercase">Прибыль</span>
                </div>
                <p className="text-2xl font-bold">11 339</p>
                <p className="text-xs text-muted-foreground">юаней</p>
              </div>
            </div>

            {/* Profitability Badge */}
            <div className="flex items-center justify-center gap-2 p-3 bg-muted rounded-lg">
              <Percent className="w-4 h-4 text-muted-foreground" />
              <span className="text-sm">Рентабельность:</span>
              <Badge variant="secondary" className="font-mono">5.54%</Badge>
            </div>

            <Separator />

            {/* Breakdown */}
            <div className="space-y-3">
              <h4 className="text-sm font-medium text-muted-foreground uppercase tracking-wider">
                Детализация
              </h4>
              
              <div className="space-y-2">
                <div className="flex justify-between text-sm">
                  <span className="text-muted-foreground">Себестоимость ед.</span>
                  <span className="font-medium">7 500 ¥</span>
                </div>
                <div className="flex justify-between text-sm">
                  <span className="text-muted-foreground">Логистика</span>
                  <span className="font-medium">150 000 ₽</span>
                </div>
                <div className="flex justify-between text-sm">
                  <span className="text-muted-foreground">НДС (15% от маржи)</span>
                  <span className="font-medium">25 050 ¥</span>
                </div>
                <div className="flex justify-between text-sm">
                  <span className="text-muted-foreground">Таможенные платежи</span>
                  <span className="font-medium">1 224 ¥</span>
                </div>
                <div className="flex justify-between text-sm">
                  <span className="text-muted-foreground">Маржа</span>
                  <span className="font-medium">541,75 ¥</span>
                </div>
              </div>
            </div>

            <Separator />

            {/* Exchange Rate */}
            <div className="flex items-center gap-2 p-3 bg-muted/50 rounded-lg">
              <Scale className="w-4 h-4 text-muted-foreground" />
              <span className="text-sm">Курс:</span>
              <span className="font-mono font-medium">1 ¥ = 14 ₽</span>
            </div>

            <Separator />

            {/* Financing Options */}
            <div className="space-y-3">
              <h4 className="text-sm font-medium text-muted-foreground uppercase tracking-wider">
                Финансирование
              </h4>
              
              <div className="space-y-2">
                <label className="flex items-center gap-3 p-3 border rounded-lg cursor-pointer hover:bg-muted/50 transition-colors">
                  <input type="checkbox" className="rounded border-input" />
                  <span className="text-sm">Кредит</span>
                </label>
                <label className="flex items-center gap-3 p-3 border rounded-lg cursor-pointer hover:bg-muted/50 transition-colors">
                  <input type="checkbox" className="rounded border-input" />
                  <span className="text-sm">Банковская гарантия</span>
                </label>
              </div>
            </div>

            {/* Additional Costs */}
            <div className="space-y-3">
              <h4 className="text-sm font-medium text-muted-foreground uppercase tracking-wider">
                Дополнительные расходы
              </h4>
              <Button variant="outline" size="sm" className="w-full bg-transparent">
                <Plus className="w-4 h-4 mr-2" />
                Добавить расход
              </Button>
            </div>
          </div>
        </ScrollArea>

        {/* Bottom Actions */}
        <div className="p-4 border-t space-y-2">
          <Button className="w-full" size="lg">
            Предварительный расчёт
          </Button>
          <Button variant="outline" className="w-full bg-transparent">
            Показать подробно
          </Button>
        </div>
      </aside>
    </div>
  )
}
