using System.Collections.Generic;

namespace MiniWord.Services
{
    /// <summary>
    /// Minimal localization service: Loc.T("key") returns the string
    /// for the currently selected language (en / ru).
    /// </summary>
    public static class Loc
    {
        public static string Lang { get; set; } = "en";

        private static readonly Dictionary<string, (string en, string ru)> S = new()
        {
            // Menus
            ["File"] = ("_File", "_Файл"),
            ["New"] = ("_New", "_Создать"),
            ["NewWindow"] = ("New _Window", "Новое _окно"),
            ["RecentFiles"] = ("Recent _Files", "Недавние _файлы"),
            ["Open"] = ("_Open...", "_Открыть..."),
            ["Save"] = ("_Save", "Со_хранить"),
            ["SaveAs"] = ("Save _As...", "Сохранить _как..."),
            ["Print"] = ("P_rint...", "_Печать..."),
            ["Preview"] = ("Print Pre_view...", "Предварительный п_росмотр..."),
            ["Exit"] = ("E_xit", "В_ыход"),
            ["Edit"] = ("_Edit", "_Правка"),
            ["Undo"] = ("_Undo", "_Отменить"),
            ["Redo"] = ("_Redo", "_Вернуть"),
            ["Cut"] = ("Cu_t", "_Вырезать"),
            ["Copy"] = ("_Copy", "_Копировать"),
            ["Paste"] = ("_Paste", "Вст_авить"),
            ["PasteText"] = ("Paste as plain te_xt", "Вставить как _текст"),
            ["SelectAll"] = ("Select _All", "Выделить _всё"),
            ["Find"] = ("_Find...", "_Найти..."),
            ["Replace"] = ("R_eplace...", "З_аменить..."),
            ["Insert"] = ("_Insert", "Вст_авка"),
            ["Image"] = ("_Image...", "_Рисунок..."),
            ["Hyperlink"] = ("_Hyperlink...", "_Гиперссылка..."),
            ["Symbol"] = ("_Symbol...", "_Символ..."),
            ["PageBreak"] = ("Page _Break", "Разрыв _страницы"),
            ["HeaderFooter"] = ("_Header && Footer...", "_Колонтитулы и нумерация..."),
            ["ExportPdf"] = ("_Export to PDF...", "_Экспорт в PDF..."),
            ["ExportHtml"] = ("Export to _HTML...", "Экспорт в _HTML..."),
            ["Tools"] = ("_Tools", "С_ервис"),
            ["Settings"] = ("_Settings...", "_Настройки..."),
            ["Help"] = ("_Help", "_Справка"),
            ["CheckUpdates"] = ("Check for _Updates...", "_Проверить обновления..."),
            ["About"] = ("_About MiniWord...", "_О программе..."),
            ["UpdateTitle"] = ("Update available", "Доступно обновление"),
            ["UpdateText"] = ("MiniWord {0} is available. You have {1}.\nUpdate now?", "Доступна версия MiniWord {0}. У вас {1}.\nОбновить сейчас?"),
            ["WhatsNew"] = ("What's new:", "Что нового:"),
            ["UpdateNow"] = ("Update now", "Обновить"),
            ["UpdateInBackground"] = ("Download in background", "Скачать в фоне"),
            ["UpdateLater"] = ("Later", "Позже"),
            ["PendingUpdateText"] = ("Update {0} has already been downloaded as you requested.\nInstall it now?", "Обновление {0} уже скачано по вашей просьбе.\nУстановить сейчас?"),
            ["RecentFilesLabel"] = ("Recent files:", "Недавние файлы:"),
            ["OpenOther"] = ("Open other...", "Открыть другой..."),
            ["NewDocument"] = ("New document", "Новый документ"),
            ["Downloading"] = ("Downloading update...", "Загрузка обновления..."),
            ["UpdateDownloadError"] = ("Failed to download the update. Please try again later.", "Не удалось загрузить обновление. Попробуйте позже."),
            ["UpToDateTitle"] = ("No updates", "Обновлений нет"),
            ["UpToDateText"] = ("You have the latest version ({0}).", "У вас установлена последняя версия ({0})."),
            ["CheckFailed"] = ("Could not check for updates. Please try again later.", "Не удалось проверить обновления. Попробуйте позже."),
            ["AboutTitle"] = ("About MiniWord", "О программе MiniWord"),
            ["AboutDescription"] = ("Lightweight free editor for .docx, .rtf, .odt and .txt", "Лёгкий бесплатный редактор .docx, .rtf, .odt и .txt"),
            ["VersionLabel"] = ("Version:", "Версия:"),
            ["ReleaseDateLabel"] = ("Release date:", "Дата релиза:"),
            ["AuthorLabel"] = ("Author:", "Автор:"),

            // Toolbar tooltips
            ["TipOpen"] = ("Open (Ctrl+O)", "Открыть (Ctrl+O)"),
            ["TipSave"] = ("Save (Ctrl+S)", "Сохранить (Ctrl+S)"),
            ["TipPrint"] = ("Print (Ctrl+P)", "Печать (Ctrl+P)"),
            ["TipUndo"] = ("Undo (Ctrl+Z)", "Отменить (Ctrl+Z)"),
            ["TipRedo"] = ("Redo (Ctrl+Y)", "Вернуть (Ctrl+Y)"),
            ["TipFont"] = ("Font", "Шрифт"),
            ["TipFontSize"] = ("Font Size", "Размер шрифта"),
            ["TipBold"] = ("Bold (Ctrl+B)", "Полужирный (Ctrl+B)"),
            ["TipItalic"] = ("Italic (Ctrl+I)", "Курсив (Ctrl+I)"),
            ["TipUnderline"] = ("Underline (Ctrl+U)", "Подчёркнутый (Ctrl+U)"),
            ["TipStrikethrough"] = ("Strikethrough", "Зачёркнутый"),
            ["TipSuperscript"] = ("Superscript", "Надстрочный"),
            ["TipSubscript"] = ("Subscript", "Подстрочный"),
            ["TipChangeCase"] = ("Change Case", "Регистр"),
            ["TipFormatPainter"] = ("Format Painter", "Формат по образцу"),
            ["TipLineSpacing"] = ("Line Spacing", "Междустрочный интервал"),
            ["TipShowMarks"] = ("Show paragraph marks", "Отображать знаки абзацев"),
            ["TipCut"] = ("Cut (Ctrl+X)", "Вырезать (Ctrl+X)"),
            ["TipCopy"] = ("Copy (Ctrl+C)", "Копировать (Ctrl+C)"),
            ["TipPaste"] = ("Paste (Ctrl+V)", "Вставить (Ctrl+V)"),
            ["CaseUpper"] = ("UPPERCASE", "ВСЕ ПРОПИСНЫЕ"),
            ["CaseLower"] = ("lowercase", "все строчные"),
            ["CaseSentence"] = ("Sentence case", "Как в предложениях"),
            ["CaseTitle"] = ("Capitalize Each Word", "Каждое Слово С Прописной"),
            ["TipAlignLeft"] = ("Align Left", "По левому краю"),
            ["TipAlignCenter"] = ("Center", "По центру"),
            ["TipAlignRight"] = ("Align Right", "По правому краю"),
            ["TipAlignJustify"] = ("Justify", "По ширине"),
            ["TipFontColor"] = ("Font Color", "Цвет текста"),
            ["TipHighlight"] = ("Text Highlight Color", "Цвет выделения текста"),
            ["Automatic"] = ("Automatic", "Авто"),
            ["NoColor"] = ("No Color", "Нет цвета"),
            ["TipBullets"] = ("Bullets", "Маркированный список"),
            ["TipNumbering"] = ("Numbering", "Нумерованный список"),
            ["TipParagraph"] = ("Paragraph Settings...", "Параметры абзаца..."),
            ["TipInsertImage"] = ("Insert Image", "Вставить рисунок"),
            ["TipPageSize"] = ("Page Size", "Размер страницы"),

            // Dialogs
            ["Paragraph"] = ("Paragraph", "Абзац"),
            ["IndentSpacing"] = ("Indentation && Spacing", "Отступы и интервалы"),
            ["LeftIndent"] = ("Left indent (pt):", "Отступ слева (пт):"),
            ["RightIndent"] = ("Right indent (pt):", "Отступ справа (пт):"),
            ["FirstLineIndent"] = ("First line indent (pt):", "Первая строка (пт):"),
            ["LineSpacing"] = ("Line spacing (pt):", "Межстрочный интервал (пт):"),
            ["SpaceBefore"] = ("Space before (pt):", "Интервал перед (пт):"),
            ["SpaceAfter"] = ("Space after (pt):", "Интервал после (пт):"),
            ["OK"] = ("OK", "ОК"),
            ["Cancel"] = ("Cancel", "Отмена"),
            ["Close"] = ("Close", "Закрыть"),

            // Find & Replace
            ["FindTitle"] = ("Find and Replace", "Найти и заменить"),
            ["FindWhat"] = ("Find what:", "Найти:"),
            ["ReplaceWith"] = ("Replace with:", "Заменить на:"),
            ["FindNext"] = ("Find Next", "Найти далее"),
            ["ReplaceBtn"] = ("Replace", "Заменить"),
            ["ReplaceAll"] = ("Replace All", "Заменить все"),
            ["MatchCase"] = ("Match case", "Учитывать регистр"),
            ["NotFound"] = ("The text was not found.", "Текст не найден."),
            ["ReplacedCount"] = ("Replaced {0} occurrence(s).", "Заменено вхождений: {0}."),

            // Hyperlink
            ["HyperlinkTitle"] = ("Insert Hyperlink", "Вставить гиперссылку"),
            ["LinkTextLabel"] = ("Text to display:", "Отображаемый текст:"),
            ["LinkUrlLabel"] = ("Address (URL):", "Адрес (URL):"),

            // Symbol
            ["SymbolTitle"] = ("Insert Symbol", "Вставить символ"),

            // Word count / autosave / PDF
            ["WordCount"] = ("Words: {0}   Characters: {1}", "Слов: {0}   Знаков: {1}"),
            ["RecoverTitle"] = ("Recover document", "Восстановление документа"),
            ["RecoverText"] = ("MiniWord found an autosaved draft that was not saved normally.\nRecover it?", "Найден черновик автосохранения, не сохранённый обычным образом.\nВосстановить его?"),
            ["FilterPdf"] = ("PDF files (*.pdf)|*.pdf", "Файлы PDF (*.pdf)|*.pdf"),
            ["PdfPrinterMissing"] = ("The \"Microsoft Print to PDF\" printer was not found in the system.", "Принтер «Microsoft Print to PDF» не найден в системе."),

            ["HFTitle"] = ("Header and Footer", "Колонтитулы"),
            ["HeaderLabel"] = ("Header text:", "Верхний колонтитул:"),
            ["FooterLabel"] = ("Footer text:", "Нижний колонтитул:"),
            ["PageNumbersCheck"] = ("Page numbers", "Нумерация страниц"),
            ["ShowOnFirstPage"] = ("Show header/footer on first page", "Колонтитул на первой странице"),
            ["PositionLabel"] = ("Number position:", "Расположение номера:"),
            ["PosFooterCenter"] = ("Bottom center", "Внизу по центру"),
            ["PosFooterRight"] = ("Bottom right", "Внизу справа"),
            ["PosHeaderRight"] = ("Top right", "Вверху справа"),
            ["PreviewTitle"] = ("Print Preview", "Предварительный просмотр"),
            ["PageStatus"] = ("Page {0} of {1}", "Страница {0} из {1}"),
            ["SettingsTitle"] = ("Settings", "Настройки"),
            ["DefaultFont"] = ("Default font:", "Шрифт по умолчанию:"),
            ["DefaultFontSize"] = ("Default font size:", "Размер шрифта по умолчанию:"),
            ["LanguageLabel"] = ("Language:", "Язык:"),
            ["LangEnglish"] = ("English", "Английский"),
            ["LangRussian"] = ("Russian", "Русский"),
            ["RestartNote"] = ("Language is applied immediately.", "Язык применяется сразу."),

            // Messages
            ["AppTitle"] = ("MiniWord", "MiniWord"),
            ["UnsavedTitle"] = ("Unsaved Changes", "Несохранённые изменения"),
            ["UnsavedText"] = ("You have unsaved changes. Save before closing?", "Есть несохранённые изменения. Сохранить?"),
            ["Error"] = ("Error", "Ошибка"),
            ["ErrorOpen"] = ("Error opening file", "Ошибка при открытии файла"),
            ["ErrorSave"] = ("Error saving file", "Ошибка при сохранении файла"),
            ["FilterDocx"] = (
                "All supported documents (*.docx;*.rtf;*.odt;*.txt)|*.docx;*.rtf;*.odt;*.txt|Word Document (*.docx)|*.docx|Rich Text Format (*.rtf)|*.rtf|OpenDocument Text (*.odt)|*.odt|Plain Text (*.txt)|*.txt|All Files (*.*)|*.*",
                "Все поддерживаемые документы (*.docx;*.rtf;*.odt;*.txt)|*.docx;*.rtf;*.odt;*.txt|Документ Word (*.docx)|*.docx|Rich Text Format (*.rtf)|*.rtf|OpenDocument Text (*.odt)|*.odt|Обычный текст (*.txt)|*.txt|Все файлы (*.*)|*.*"),
            ["FilterDocxSave"] = (
                "Word Document (*.docx)|*.docx|Rich Text Format (*.rtf)|*.rtf|OpenDocument Text (*.odt)|*.odt|Plain Text (*.txt)|*.txt",
                "Документ Word (*.docx)|*.docx|Rich Text Format (*.rtf)|*.rtf|OpenDocument Text (*.odt)|*.odt|Обычный текст (*.txt)|*.txt"),
            ["FilterHtml"] = ("HTML Page (*.html)|*.html", "HTML-страница (*.html)|*.html"),
            ["FilterImages"] = ("Images (*.png;*.jpg;*.jpeg;*.bmp;*.gif)|*.png;*.jpg;*.jpeg;*.bmp;*.gif", "Изображения (*.png;*.jpg;*.jpeg;*.bmp;*.gif)|*.png;*.jpg;*.jpeg;*.bmp;*.gif"),
            ["DefaultDocName"] = ("Document", "Документ"),
        };

        public static string T(string key)
        {
            if (S.TryGetValue(key, out var pair))
                return Lang == "ru" ? pair.ru : pair.en;
            return key;
        }
    }
}
