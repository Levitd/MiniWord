# MiniWord

Лёгкий бесплатный текстовый редактор для Windows с поддержкой формата **.docx** (совместим с MS Word). Минимальный набор функций без перегруженного интерфейса — альтернатива тяжёлым офисным пакетам для простых документов.

*A minimal, free Windows word processor with native **.docx** support — a lightweight alternative to full office suites for simple documents.*

![Platform](https://img.shields.io/badge/platform-Windows-blue) ![.NET](https://img.shields.io/badge/.NET-6.0-purple) ![License](https://img.shields.io/badge/license-MIT-green)

## Возможности / Features

- 📄 Открытие и сохранение **.docx** (DocumentFormat.OpenXml — официальный SDK Microsoft)
- 🔤 Шрифт, размер (в пунктах, как в Word), **жирный** / *курсив* / подчёркнутый
- 🎨 Цвет текста и заливка (выделение маркером)
- 📐 Выравнивание, отступы, межстрочные интервалы, отступ первой строки
- 📝 Маркированные и нумерованные списки
- 🖼️ Вставка картинок (сохраняются в .docx)
- 📃 Вид страницы (белый лист, как в Word), размеры A4 / Letter / A5, границы страниц
- 🔢 Колонтитулы и нумерация страниц (сохраняются в .docx как настоящие header/footer с полем PAGE)
- 🖨️ Печать и предварительный просмотр
- ↩️ Undo / Redo
- 🌍 Интерфейс на русском и английском (переключение в настройках)

## Скачать / Download

**[⬇ Скачать установщик (Releases)](https://github.com/Levitd/MiniWord/releases/latest)** — `MiniWordSetup-x.x.x.exe` (~3 МБ)

Установщик на русском и английском: выбор папки установки, ярлыки на рабочем столе и в меню Пуск (из меню Пуск можно закрепить на панели задач: ПКМ → «Закрепить на панели задач»).

Требуется [.NET 6 Desktop Runtime](https://dotnet.microsoft.com/download/dotnet/6.0/runtime) (или новее) — если он не установлен, установщик предложит открыть страницу загрузки.

*Requires .NET 6 Desktop Runtime or newer — the installer checks for it and offers the download page if missing.*

## Сборка / Build

Требуется [.NET 6 SDK](https://dotnet.microsoft.com/download/dotnet/6.0) (или новее) на Windows.

```bash
cd MiniWord
dotnet build
dotnet run
```

Портативная сборка одним exe:

```bash
dotnet publish -c Release -r win-x64 --self-contained -p:PublishSingleFile=true
```

## Технологии / Tech stack

- **C# / WPF** (.NET 6, `net6.0-windows`)
- `RichTextBox` + `FlowDocument` — модель документа и редактирование
- **DocumentFormat.OpenXml** — чтение/запись WordprocessingML (.docx)
- WPF `DocumentPaginator` + XPS — печать и предпросмотр

## Ограничения / Limitations

Сознательно вне рамок проекта: таблицы, стили, сноски, рецензирование, совместное редактирование. Разбивка на страницы в окне редактора — приближённая (точная — в предпросмотре печати).

## Лицензия / License

MIT
