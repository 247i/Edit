var module_name = "Sort Array Line";
var module_ver = "1.00";
var module_author = "DJ N-4ceR / http://www.n4cer.de";

function FormatCode1()
{
  var ed  = newEditor();
  ed.assignActiveEditor();
  var input = ed.selText();
  var words = input.split(' ');
  words.sort();
  var output = words.join(' ');
  ed.selText(output);
}

function FormatCode2()
{
  var ed  = newEditor();
  ed.assignActiveEditor();
  var input = ed.selText();
  var words = input.split(', ');
  words.sort();
  var output = words.join(', ');
  ed.selText(output);
}

function Init()
{
  addMenuItem("Sort @array=qw(word1 word2 word3 ...)", "Sort Array Line", "FormatCode1");
  addMenuItem("Sort @array=('word1', 'word2', 'word3', ...)", "Sort Array Line", "FormatCode2");
  addMenuItem("Sort @array=(\"word1\", \"word2\", \"word3\", ...)", "Sort Array Line", "FormatCode2");
}

