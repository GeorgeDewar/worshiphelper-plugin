using System;
using System.Linq;

namespace PowerWorshipVSTO
{
    class ScriptureReference
    {
        String bookName;
        int chapterNum;
        int verseNumStart;
        int verseNumEnd;
        
        public static ScriptureReference parse(Bible bible, String bookName, String reference)
        {
            var scriptureReference = new ScriptureReference();

            var book = bible.books.Find(bookItem => bookItem.name.ToLower() == bookName.ToLower());
            if (book == null) throw new Exception("Book does not exist");

            var referenceParts = reference.Split(new char[] { ':', '-' });

            scriptureReference.bookName = book.name;
            scriptureReference.chapterNum = Int32.Parse(referenceParts[0]);
            var chapter = book.chapters.Find(chapterItem => chapterItem.number == scriptureReference.chapterNum);
            if (chapter == null) throw new Exception("Chapter does not exist");

            if (referenceParts.Length > 2)
            {
                scriptureReference.verseNumStart = Int32.Parse(referenceParts[1]);
                scriptureReference.verseNumEnd = Int32.Parse(referenceParts[2]);
            }
            else if (referenceParts.Length > 1)
            {
                scriptureReference.verseNumStart = Int32.Parse(referenceParts[1]);
                scriptureReference.verseNumEnd = scriptureReference.verseNumStart;
            }
            else
            {
                // No verses were specified, so use the whole range
                scriptureReference.verseNumStart = 1;
                scriptureReference.verseNumEnd = chapter.verses.Last().number;
            }

            if (scriptureReference.verseNumEnd < scriptureReference.verseNumStart) throw new Exception("Verse range end is before start");
            if (scriptureReference.verseNumStart < chapter.verses.First().number)  throw new Exception("Verse range is before beginning of chapter");
            if (scriptureReference.verseNumEnd > chapter.verses.Last().number) throw new Exception("Verse range goes past end of chapter");

            return scriptureReference;
        }
    }
}
