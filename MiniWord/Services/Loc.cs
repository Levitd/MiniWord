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
            ["Open"] = ("_Open...", "_Открыть..."),
            ["Save"] = ("_Save", "Со_хранить"),
            ["SaveAs"] = ("Save _As...", "Сохранить _как..."),
            ["Print"] = ("P_rint...", "_Печать..."),
            ["Preview"] = ("Print Pre_view...", "Предварительный п_росмотр..."),
            ["Exit"] = ("E_xit", "В_ыход"),
            ["Edit"] = ("_Edit", "_Правка"),
            ["Undo"] = ("_Undo", "_Отменить"),
            ["Redo"] = ("_Redo", "_Вернуть"),
            ["Insert"] = ("_Insert", "Вст_авка"),
            ["Image"] = ("_Image...", "_Рисунок..."),
            ["HeaderFooter"] = ("_Header && Footer...", "_Колонтитулы и нумерация..."),
            ["Tools"] = ("_Tools", "С_ервис"),
            ["Settings"] = ("_Settings...", "_Настройки..."),

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

            ["HFTitle"] = ("Header and Footer", "Колонтитулы"),
            ["HeaderLabel"] = ("Header text:", "Верхний колонтитул:"),
            ["FooterLabel"] = ("Footer text:", "Нижний колонтитул:"),
            ["PageNumbersCheck"] = ("Page numbers", "Нумерация страниц"),
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
            ["FilterDocx"] = ("Word Documents (*.docx)|*.docx|All Files (*.*)|*.*", "Документы Word (*.docx)|*.docx|Все файлы (*.*)|*.*"),
            ["FilterDocxSave"] = ("Word Documents (*.docx)|*.docx", "Документы Word (*.docx)|*.docx"),
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
