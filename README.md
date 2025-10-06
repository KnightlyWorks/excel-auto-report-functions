# 📊 VBA Excel Tools(simple) by KnightlyWorks

**Author:** KnightlyWorks  
**License:** MIT  
**Version:** 1.0

---

## 🌍 Languages / Jazyky / Языки

- [🇬🇧 English](#-english)
- [🇸🇰 Slovenčina](#-slovenčina)
- [🇷🇺 Русский](#-русский)

---

## 🇬🇧 English

### 📝 Description

A collection of useful VBA macros for Excel automation. Currently includes tools for image insertion and automatic cell formatting.

### 📦 Modules

#### 1. `InsertImage.bas`
Inserts and stretches images into selected cell ranges.

**Features:**
- Select any cell range and insert an image
- Image automatically stretches to fill the entire selected area
- No aspect ratio preservation (fills completely)
- Automatic cleanup of old images in the target area
- Customizable default folder path

**Installation:**
1. Press `Alt+F11` to open VBA Editor
2. Go to `File → Import File`
3. Select `InsertImage.bas`
4. Press `Alt+Q` to exit

**Usage:**
- Select cell range where you want the image
- Press `Alt+F8` → Select `InsertImageFitToSelection` → Run
- Choose image from dialog
- Image will be inserted and stretched to fit

**Hotkey (Optional):**
1. Press `Alt+F8`
2. Select `InsertImageFitToSelection`
3. Click `Options`
4. Set a letter (e.g., `I`) → Creates `Ctrl+Shift+I` hotkey

---

#### 2. `GreenOnOK.bas`
Automatically colors cells green when "OK" is entered.

**Features:**
- Monitors specific column for text input
- Automatically applies green formatting when trigger text is entered
- Customizable trigger text, target column, and colors
- Case-insensitive matching (optional)
- Resets formatting when text is changed

**Installation:**
1. Press `Alt+F11` to open VBA Editor
2. **Double-click** the target sheet in the left panel (e.g., `Sheet1`)
3. Paste the entire code from `GreenOnOK.bas`
4. Press `Alt+Q` to exit

⚠️ **Important:** This code must be placed in the **Sheet Module**, not a regular module!

**Usage:**
- Type "OK" in column A
- Cell automatically turns green with white bold text
- Change the text → formatting resets automatically

---

### ⚙️ Configuration

Both modules have configuration constants at the top of the file:

**InsertImage.bas:**
```vba
Const FOLDER_PATH = "C:\Users\Pictures\"      ' Default image folder
Const ALLOW_MULTIPLE_IMAGES = False           ' Keep old images?
Const IS_DEBUG = False                        ' Enable debug logging
```

**GreenOnOK.bas:**
```vba
Const TRIGGER_TEXT = "OK"                     ' Text that triggers formatting
Const TARGET_COLUMN = "A"                     ' Column to monitor
Const GREEN_COLOR = 5287936                   ' RGB color code
Const CASE_SENSITIVE = False                  ' Case-sensitive matching
Const IS_DEBUG = False                        ' Enable debug logging
```

---

### 🐛 Debug Mode

Set `IS_DEBUG = True` in any module to enable logging to Immediate Window (`Ctrl+G` in VBA Editor).

---

### 📄 Requirements

- Microsoft Excel 2010 or later
- Macros must be enabled
- File must be saved as `.xlsm` (Excel Macro-Enabled Workbook)

---

## 🇸🇰 Slovenčina

### 📝 Popis

Zbierka užitočných VBA makier pre automatizáciu práce v Exceli. Momentálne obsahuje nástroje na vkladanie obrázkov a automatické formátovanie buniek.

### 📦 Moduly

#### 1. `InsertImage.bas`

Vkladá a roztahuje obrázky do vybraných oblastí buniek.

**Funkcie:**

* Výber ľubovoľnej oblasti buniek na vloženie obrázka
* Automatické roztiahnutie obrázka na celú vybranú oblasť
* Bez zachovania pomerov strán (vyplní úplne)
* Automatické vyčistenie starých obrázkov v cieľovej oblasti
* Prispôsobiteľná predvolená cesta k priečinku

**Inštalácia:**

1. Stlačte `Alt+F11` pre otvorenie VBA editora
2. Prejdite na `Súbor → Importovať súbor`
3. Vyberte `InsertImage.bas`
4. Stlačte `Alt+Q` pre ukončenie

**Použitie:**

* Vyberte oblasť buniek, do ktorej chcete vložiť obrázok
* Stlačte `Alt+F8` → Vyberte `InsertImageFitToSelection` → Spustiť
* Vyberte obrázok v dialógovom okne
* Obrázok bude vložený a roztiahnutý podľa veľkosti buniek

**Klávesová skratka (voliteľne):**

1. Stlačte `Alt+F8`
2. Vyberte `InsertImageFitToSelection`
3. Kliknite na `Možnosti`
4. Nastavte písmeno (napr. `I`) → Vytvorí skratku `Ctrl+Shift+I`

---

#### 2. `GreenOnOK.bas`

Automaticky zafarbuje bunky na zeleno pri zadaní "OK".

**Funkcie:**

* Monitoruje konkrétny stĺpec na zadanie textu
* Automaticky aplikuje zelené formátovanie pri zadaní spúšťacieho textu
* Prispôsobiteľný aktivujúci text, cieľový stĺpec a farby
* Necitlivosť na veľkosť písmen (voliteľné)
* Obnovenie formátovania pri zmene textu

**Inštalácia:**

1. Stlačte `Alt+F11` pre otvorenie VBA editora
2. **Dvakrát kliknite** na požadovaný hárok v ľavom paneli (napr. `Hárok1`)
3. Vložte celý kód z `GreenOnOK.bas`
4. Stlačte `Alt+Q` pre ukončenie

⚠️ **Dôležité:** Tento kód musí byť umiestnený v **module hárka**, nie v bežnom module!

**Použitie:**

* Napíšte "OK" v stĺpci A
* Bunka sa automaticky zafarbí na zeleno, s bielym tučným textom
* Zmeňte text → formátovanie sa automaticky obnoví

---

### ⚙️ Konfigurácia

Oba moduly majú konfiguračné konštanty na začiatku súboru:

**InsertImage.bas:**

```vba
Const FOLDER_PATH = "C:\Users\Pictures\"      ' Predvolený priečinok s obrázkami
Const ALLOW_MULTIPLE_IMAGES = False           ' Zachovať staré obrázky?
Const IS_DEBUG = False                        ' Povoliť logovacie hlásenia
```

**GreenOnOK.bas:**

```vba
Const TRIGGER_TEXT = "OK"                     ' Text spúšťajúci formátovanie
Const TARGET_COLUMN = "A"                     ' Monitorovaný stĺpec
Const GREEN_COLOR = 5287936                   ' RGB kód farby
Const CASE_SENSITIVE = False                  ' Rozlišovať veľkosť písmen
Const IS_DEBUG = False                        ' Povoliť logovacie hlásenia
```

---

### 🐛 Režim ladenia

Nastavte `IS_DEBUG = True` v ľubovoľnom module pre povolenie logovania do okna Immediate (`Ctrl+G` vo VBA editore).

---

### 📄 Požiadavky

* Microsoft Excel 2010 alebo novší
* Makrá musia byť povolené
* Súbor musí byť uložený ako `.xlsm` (Zošit Excelu s podporou makier)


## 🇷🇺 Русский

### 📝 Описание

Коллекция полезных VBA макросов для автоматизации работы в Excel. В данный момент включает инструменты для вставки изображений и автоматического форматирования ячеек.

### 📦 Модули

#### 1. `InsertImage.bas`
Вставляет и растягивает изображения в выбранные диапазоны ячеек.

**Возможности:**
- Выбор любого диапазона ячеек для вставки изображения
- Автоматическое растягивание изображения на всю выбранную область
- Без сохранения пропорций (заполняет полностью)
- Автоматическая очистка старых изображений в целевой области
- Настраиваемый путь к папке по умолчанию

**Установка:**
1. Нажмите `Alt+F11` для открытия редактора VBA
2. Перейдите в `Файл → Импорт файла`
3. Выберите `InsertImage.bas`
4. Нажмите `Alt+Q` для выхода

**Использование:**
- Выделите диапазон ячеек, куда нужно вставить изображение
- Нажмите `Alt+F8` → Выберите `InsertImageFitToSelection` → Выполнить
- Выберите изображение в диалоговом окне
- Изображение будет вставлено и растянуто по размеру ячеек

**Горячая клавиша (опционально):**
1. Нажмите `Alt+F8`
2. Выберите `InsertImageFitToSelection`
3. Нажмите `Параметры`
4. Задайте букву (например, `I`) → Создаст горячую клавишу `Ctrl+Shift+I`

---

#### 2. `GreenOnOK.bas`
Автоматически окрашивает ячейки в зелёный цвет при вводе "OK".

**Возможности:**
- Отслеживает указанный столбец на предмет ввода текста
- Автоматически применяет зелёное форматирование при вводе триггерного текста
- Настраиваемый текст-триггер, целевой столбец и цвета
- Нечувствительность к регистру (опционально)
- Сброс форматирования при изменении текста

**Установка:**
1. Нажмите `Alt+F11` для открытия редактора VBA
2. **Дважды кликните** по нужному листу на левой панели (например, `Лист1`)
3. Вставьте весь код из `GreenOnOK.bas`
4. Нажмите `Alt+Q` для выхода

⚠️ **Важно:** Этот код должен быть размещён в **модуле листа**, а не в обычном модуле!

**Использование:**
- Введите "OK" в столбце A
- Ячейка автоматически окрасится в зелёный с белым жирным текстом
- Измените текст → форматирование автоматически сбросится

---

### ⚙️ Конфигурация

Оба модуля имеют настраиваемые константы в начале файла:

**InsertImage.bas:**
```vba
Const FOLDER_PATH = "C:\Users\Pictures\"      ' Папка с изображениями по умолчанию
Const ALLOW_MULTIPLE_IMAGES = False           ' Сохранять старые изображения?
Const IS_DEBUG = False                        ' Включить логирование отладки
```

**GreenOnOK.bas:**
```vba
Const TRIGGER_TEXT = "OK"                     ' Текст, запускающий форматирование
Const TARGET_COLUMN = "A"                     ' Отслеживаемый столбец
Const GREEN_COLOR = 5287936                   ' Код цвета RGB
Const CASE_SENSITIVE = False                  ' Учитывать регистр
Const IS_DEBUG = False                        ' Включить логирование отладки
```

---

### 🐛 Режим отладки

Установите `IS_DEBUG = True` в любом модуле для включения логирования в окно Immediate (`Ctrl+G` в редакторе VBA).

---

### 📄 Требования

- Microsoft Excel 2010 или новее
- Макросы должны быть включены
- Файл должен быть сохранён как `.xlsm` (Книга Excel с поддержкой макросов)

---

## 📜 License

MIT License - feel free to use and modify!

---

**Made with ❤️ by KnightlyWorks**