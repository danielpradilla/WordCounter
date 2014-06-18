using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using Microsoft.Office.Interop.Word;

namespace WordCounter
{
    class Counter
    {
        /*
         * Given a filename (full path to file), returns the MS Word word count
         * Author: Daniel Pradilla info@danielpradilla.info
         */
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public CounterData getCounters(String filePath)
        {

            CounterData counters = new CounterData { wordCount = 0, charCount = 0, paraCount = 0, pageCount = 0 };
            string fileName = filePath.Substring(filePath.LastIndexOf("\\") + 1);
            log.Info("Counting " + fileName);
            Microsoft.Office.Interop.Word.Application word = null;
            Microsoft.Office.Interop.Word.Document doc = null;

            object missing = Type.Missing;
            object saveChanges = false;
            object includeFootnotesAndEndnotes = true;
            Microsoft.Office.Interop.Word.WdStatistic wordStats = Microsoft.Office.Interop.Word.WdStatistic.wdStatisticWords;
            Microsoft.Office.Interop.Word.WdStatistic charStats = Microsoft.Office.Interop.Word.WdStatistic.wdStatisticCharacters;
            Microsoft.Office.Interop.Word.WdStatistic paraStats = Microsoft.Office.Interop.Word.WdStatistic.wdStatisticParagraphs;
            Microsoft.Office.Interop.Word.WdStatistic pageStats = Microsoft.Office.Interop.Word.WdStatistic.wdStatisticPages;

            try
            {

                word = new Microsoft.Office.Interop.Word.Application();
                doc = new Microsoft.Office.Interop.Word.Document();
                object objFilePath = @filePath;

                /*
                 Document OpenNoRepairDialog(ref Object FileName,
	                    ref Object ConfirmConversions, ref Object ReadOnly, ref Object AddToRecentFiles, ref Object PasswordDocument, 
                        ref Object PasswordTemplate, ref Object Revert, ref Object WritePasswordDocument, ref Object WritePasswordTemplate,
	                    ref Object Format, ref Object Encoding, ref Object Visible, ref Object OpenAndRepair,
	                    ref Object DocumentDirection, ref Object NoEncodingDialog, ref Object XMLTransform
                    )
                 */
                doc = word.Documents.OpenNoRepairDialog(ref objFilePath,
                    ref missing, true, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing);
                if (doc !=null)
                {
                    doc.Activate();

                    counters.wordCount = doc.ComputeStatistics(wordStats, ref includeFootnotesAndEndnotes);
                    counters.charCount = doc.ComputeStatistics(charStats, ref includeFootnotesAndEndnotes);
                    counters.paraCount = doc.ComputeStatistics(paraStats, ref includeFootnotesAndEndnotes);
                    counters.pageCount = doc.ComputeStatistics(pageStats, ref includeFootnotesAndEndnotes);
                    log.Info(counters);
                }
                //doc.Save();
                word.Quit(ref saveChanges, ref missing, ref missing);
            }
            catch (Exception ex)
            {
                word.Quit(ref saveChanges, ref missing, ref missing);
                log.Error(ex);
            }

            return counters;

        }

    }


    class CounterData
    {
        public int charCount { get; set; }
        public int wordCount { get; set; }
        public int paraCount { get; set; }
        public int pageCount { get; set; }

    }


}
