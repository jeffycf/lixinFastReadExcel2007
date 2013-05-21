lixinFastReadExcel07
====================

<h1>一个快速读取excel2007内容的类库</h1>
<p>该类库实现读取excel中sheet的名字，及sheet的内容功能。对读取的内容为没有任何样式的字符串，所以速度比较快。</p>
<h2>原理</h2>
<p>通过解压excel的zip包，读取xl/workbook.xml 文件，获取excel所有的sheet的名称。读取xl/worksheets/sheet*.xml 文件获得sheet的内容</p>

