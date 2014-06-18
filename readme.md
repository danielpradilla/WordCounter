##WordCounter

These are a couple of classes for counting words in Microsoft Word Documents. One in groovy and one in C#

The groovy one I took from [https://gist.github.com/kaisternad/7736686](https://gist.github.com/kaisternad/7736686)
It relies on [Apache Tika](http://tika.apache.org/) 

I realized that sometimes Tika reports the last saved word count stored in the metadata of the file. This word count may or may not be the real word count of the document! 

So if you're going to stick the groovy script and Apache Tika you've been warned.

###A real(er) word count
For a real word count you need to open the word document, force a re-count and get that number. That's what the C# class does. It relies on [Microsoft.Office.Interop.Word](http://msdn.microsoft.com/en-us/library/microsoft.office.interop.word.aspx).

Needless to say, you need Windows MS Word installed in the machine executing the class.


###Usage
Drop them in your project, use them as you want and drop me a note :)

