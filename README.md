This is a side project from [Antragsgrün](https://github.com/CatoTH/antragsgruen).

A demo script using the default template is:

```<?php

require_once(__DIR__ . DIRECTORY_SEPARATOR . 'vendor/autoload.php');

$html = '<p>This is a demo for the converter.</p>
<p>The converter supports the following styles:</p>
<ul>
    <li>Lists (UL / OL)</li>
    <li><strong>STRONG</strong></li>
    <li><u>U</u> (underlined)</li>
    <li><s>S</s> (strike-through)</li>
    <li><em>EM</em> (emphasis / italic)</li>
    <li>Line<br>breaks with BR</li>
</ul>
<blockquote>You can also use BLOCKQUOTE, though it lacks specific styling for now</blockquote>';

$html2 = '<p>You might be interested<br>in the fact that this converter<br>also supports<br>line numbering<br>for selected paragraphs</p>
<p>Dummy Line<br>Dummy Line<br>Dummy Line<br>Dummy Line<br>Dummy Line</p>';

try {
    $text = new \CatoTH\HTML2OpenDocument\Text();
    $text->addHtmlTextBlock('<h1>Test Page</h1>');
    $text->addHtmlTextBlock($html, false);
    $text->addHtmlTextBlock('<h2>Line Numbering</h2>');
    $text->addHtmlTextBlock($html2, true);
    echo $text->finishAndGetOdt();
} catch (\Exception $e) {
    var_dump($e);
}
?>
```