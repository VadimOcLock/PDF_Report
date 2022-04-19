﻿using iText.Html2pdf;
using System.Text;

namespace PDF_Report
{
    public class AdmissionReport
    {
        private static readonly string _htmlTemplate = "<!DOCTYPE html><html ><head> <meta charset='UTF-8'> <title>Протокол допуска</title></head><style type='text/css'>body{font-family: Arial, sans-serif; font-size: 12px; font-weight: normal; line-height: 1.3; margin-top: 2cm; margin-bottom: 2cm; margin-left: 3cm; margin-right: 1cm;}.right-upper{text-align: right; position: absolute; top: 0px; right: 0px; font-size: 14px;}.hl{text-align: center; font-size: 18px;}.tbl{border-collapse: collapse; text-align: center; border-spacing: 0; width: 100%;}.style0{margin-bottom: 20px; border-collapse: collapse; border-spacing: 0; width: 100%}.tbl td{border-color: black; border-style: solid; border-width: 1px; font-family: Arial, sans-serif; font-size: 14px; overflow: hidden; padding: 10px 5px; word-break: normal;}@page{size: A4 portrait; /* set page margins */ margin: 0.5cm; @top-center{content: element(header);}@bottom-center{content: element(footer);}}</style><body> <table> <thead> <tr> <div class='right-upper'> [ABB]/[YEAR]-[ION_NAME]-[PRTCL_NUM] </div><div class='hl'> <p>Протокол <u>№ [PRTCL_NUM]</u> от <u>[CREATE_DATA]</u></p></div><div class='hl'> <p> Определения неоднородности флюенса ионов <sup>[ION_NUM]</sup>[ION_NAME] с энергией [ENERGY_VALUE] МэВ/N на испытательном стенде [BENCH]</p></div></tr></thead> <tbody> <tr> <p><b>1. Цель:</b> [TARGET].</p></tr><tr> <p><b>2. Время и место определения неоднородности флюенса ионов:</b></p><p>Проводилась в период с : [START_DATA] [START_TIME] по [END_DATA] [END_TIME] в [LOCATION].</p></tr><tr> <p><b>3. Условия определения неоднородности флюенса ионов:</b></p><ul> <li>температура окружающей среды: [TEMP] °С;</li><li>атмосферное давление: [TENSION] мм рт.ст.;</li><li>относительная влажность воздуха: [WET] %.</li></ul> </tr><tr> <p><b>4. Средства определения неоднородности флюенса ионов:</b></p><ul> <li>испытательный стенд: [BENCH]</li><li>трековые мембраны (лавсановая плёнка);</li><li>установка для травления лавсановой плёнки;</li><li>растровый электронный микроскоп ТM-3000 (Hitachi, Япония);</li><li>cистема оцифровки видеосигнала «GALLERY-512».</li></ul> </tr><tr> <p><b>5. Методика определения неоднородности флюенса ионов.</b></p><p>5.1. Проводилась в соответствии с «Методикой измерений флюенса тяжелых заряженных частиц с помощью трековых мембран на основе лавсановой пленки» ЦДКТ1.027.012-2015.</p></tr><tr> <p><b>6. Результаты определения неоднородности флюенса ионов</b> <sup>[ION_NUM]</sup>[ION_NAME] представлены в таблице 1:</p><table class='style0'> <tr> <td>N=[INTENSE] c <sup><small>-1</small></sup></td><td>Φ=[COUNT] частиц*см<sup><small>-2</small></sup></td></tr></table> <table class='tbl'> <tr> <tr> <td>ТД1</td><td>ТД2</td><td>ТД3</td><td>ТД4</td><td>ТД5</td></tr><tr> <td>[TD_1]</td><td>[TD_2]</td><td>[TD_3]</td><td>[TD_4]</td><td>[TD_5]</td></tr><tr> <td>ТД6</td><td>ТД7</td><td>ТД8</td><td>ТД9</td><td>Среднее зн.</td></tr><tr> <td>[TD_6]</td><td>[TD_7]</td><td>[TD_8]</td><td>[TD_9]</td><td>[AVG_TD_VAL]</td></tr></tr></table> <p>Коэффициент : К <sub>расчетный</sub>=[K_VALUE] ± [K_MEASURE]</p><p>Неоднородность флюенса ионов составила : <u>[GENEITY]</u></p></tr><tr> <p><b>7. Принято решение о продолжении работ на ионе / повторной настройке пучка</b></p><p>в [RESULT_TIME]</p></tr><tr> <table> <tr> <td style='text-align:center; padding-right: 10px;';>Ответственный за проведение испытаний в испытательную смену от ООО \"НПП \"Детектор\"</td><td style='text-align:center; padding-left: 10px;'>Ответственный за проверку от ЛЯР ОИЯИ</td></tr><tr> <td style='width: 50%; text-align: center; padding-top: 10px; padding-right: 10px;'>_____________________ (&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;)</td><td style='width: 50%; text-align: center; padding-top: 10px; padding-left: 10px;'>_____________________ (&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;)</td></tr></table> </tr></tbody> </table></body></html>";

        private StringBuilder _htmlDefault = new StringBuilder(_htmlTemplate);
        private StringBuilder _html = new StringBuilder(_htmlTemplate);

        Dictionary<string, string> ParamsReplace = new Dictionary<string, string>()
        {
            {"[ABB]", "0" },
            {"[YEAR]", "0" },
            {"[ION_NAME]", "0" },
            {"[PRTCL_NUM]", "0" },
            {"[CREATE_DATA]", "0" },
            {"[ION_NUM]", "0" },
            {"[ENERGY_VALUE]", "0" },
            {"[BENCH]", "0" },
            {"[TARGET]", "0" },
            {"[START_DATA]", "0" },
            {"[START_TIME]", "0" },
            {"[END_DATA]", "0" },
            {"[END_TIME]", "0" },
            {"[LOCATION]", "0" },
            {"[TEMP]", "0" },
            {"[TENSION]", "0" },
            {"[WET]", "0" },
            {"[INTENSE]", "0" },
            {"[COUNT]", "0" },
            {"[TD_1]", "0" },
            {"[TD_2]", "0" },
            {"[TD_3]", "0" },
            {"[TD_4]", "0" },
            {"[TD_5]", "0" },
            {"[TD_6]", "0" },
            {"[TD_7]", "0" },
            {"[TD_8]", "0" },
            {"[TD_9]", "0" },
            {"[AVG_TD_VAL]", "0" },
            {"[K_VALUE]", "0" },
            {"[K_MEASURE]", "0" },
            {"[GENEITY]", "0" },
            {"[RESULT_TIME]", "0" }
        };

        public AdmissionReport()
        {
            _htmlDefault = Rename(_htmlDefault);
        }

        private StringBuilder Rename(StringBuilder html)
        {
            foreach (var p in ParamsReplace)
            {
                html.Replace(p.Key, p.Value);
            }

            return html;
        }



        /// <summary>
        /// Возвращает преобразованный html протокол.<br/>
        /// Если поля не были проинициализированны, то они будут заполенены значениями по умолчанию.
        /// </summary>
        public string GetHtml()
        {
            _html = Rename(_html);

            return _html.ToString();
        }

        public void ConvertToPDF(string xPath)
        {
            using (FileStream fs = new FileStream(xPath, FileMode.OpenOrCreate))
            {
                HtmlConverter.ConvertToPdf(GetHtml(), fs);
            }
        }

        /// <summary>
        /// Возвращает непреобразованный шаблон html протокола.
        /// </summary>
        public string GetHtmlTemplate() => _htmlTemplate;
    }
}
