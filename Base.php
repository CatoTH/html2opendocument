<?php

/**
 * @link https://github.com/CatoTH/html2opendocument
 * @author Tobias Hößl <tobias@hoessl.eu>
 * @license https://opensource.org/licenses/MIT
 */

namespace CatoTH\HTML2OpenDocument;

abstract class Base
{
    const NS_OFFICE   = 'urn:oasis:names:tc:opendocument:xmlns:office:1.0';
    const NS_TEXT     = 'urn:oasis:names:tc:opendocument:xmlns:text:1.0';
    const NS_FO       = 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0';
    const NS_DRAW     = 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0';
    const NS_STYLE    = 'urn:oasis:names:tc:opendocument:xmlns:style:1.0';
    const NS_SVG      = 'urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0';
    const NS_TABLE    = 'urn:oasis:names:tc:opendocument:xmlns:table:1.0';
    const NS_CALCTEXT = 'urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0';
    const NS_XLINK    = 'http://www.w3.org/1999/xlink';


    /** @var \DOMDocument|null */
    protected $doc = null;

    /** @var bool */
    protected $DEBUG = false;
    protected $trustHtml = false;

    /** @var string */
    protected $tmpPath = '/tmp/';

    /** @var \ZipArchive */
    private $zip;

    /** @var @string */
    private $tmpZipFile;

    /** @var null|string */
    protected $pageWidth        = null;
    protected $pageHeight       = null;
    protected $printOrientation = null;
    protected $marginTop        = null;
    protected $marginLeft       = null;
    protected $marginRight      = null;
    protected $marginBottom     = null;

    /** @var string */
    protected $colorIns         = '#008800';
    protected $colorDel         = '#880000';

    /**
     * @throws \Exception
     */
    public function __construct(string $templateFile, array $options = [])
    {
        $template = file_get_contents($templateFile);
        if (isset($options['tmpPath']) && $options['tmpPath'] !== '') {
            $this->tmpPath = $options['tmpPath'];
        }
        if (isset($options['trustHtml'])) {
            $this->trustHtml = ($options['trustHtml'] === true);
        }
        if (isset($options['colorIns'])) {
            $this->colorIns = $options['colorIns'];
        }

        if(!file_exists($this->tmpPath)){
            mkdir($this->tmpPath);
        }

        $this->tmpZipFile = $this->tmpPath . uniqid('zip-');
        file_put_contents($this->tmpZipFile, $template);

        $this->zip = new \ZipArchive();
        if ($this->zip->open($this->tmpZipFile) !== true) {
            throw new \Exception("cannot open <$this->tmpZipFile>\n");
        }

        $content = $this->zip->getFromName('content.xml');

        $this->doc = new \DOMDocument();
        $this->doc->loadXML($content);

    }

    public function finishAndGetDocument(): string
    {
        $content = $this->create();

        $this->zip->deleteName('content.xml');
        $this->zip->addFromString('content.xml', $content);

        $this->writePageStyles();

        $this->zip->close();

        $content = file_get_contents($this->tmpZipFile);
        unlink($this->tmpZipFile);
        return $content;
    }

    protected function writePageStyles(): void
    {
        $stylesStr = $this->zip->getFromName('styles.xml');
        $styles = new \DOMDocument();
        $styles->loadXML($stylesStr);

        foreach ($styles->getElementsByTagNameNS(static::NS_STYLE, 'page-layout-properties') as $element) {
            /** @var \DOMElement $element */
            if ($this->pageWidth) {
                $element->setAttribute('fo:page-width', $this->pageWidth);
            }
            if ($this->pageHeight) {
                $element->setAttribute('fo:page-height', $this->pageHeight);
            }
            if ($this->printOrientation) {
                $element->setAttribute('style:print-orientation', $this->printOrientation);
            }
            if ($this->marginTop) {
                $element->setAttribute('fo:margin-top', $this->marginTop);
            }
            if ($this->marginLeft) {
                $element->setAttribute('fo:margin-left', $this->marginLeft);
            }
            if ($this->marginRight) {
                $element->setAttribute('fo:margin-right', $this->marginRight);
            }
            if ($this->marginBottom) {
                $element->setAttribute('fo:margin-bottom', $this->marginBottom);
            }
        }

        $xml = $styles->saveXML();

        $rows = explode("\n", $xml);
        $rows[0] .= "\n";
        $stylesStr = implode('', $rows) . "\n";

        $this->zip->deleteName('styles.xml');
        $this->zip->addFromString('styles.xml', $stylesStr);
    }

    abstract function create(): string;

    protected function purifyHTML(string $html, array $config): string
    {
        $configInstance               = \HTMLPurifier_Config::create($config);
        $configInstance->autoFinalize = false;
        $purifier                     = \HTMLPurifier::instance($configInstance);
        $purifier->config->set('Cache.SerializerPath', $this->tmpPath);

        return $purifier->purify($html);
    }

    public function html2DOM(string $html): \DOMNode
    {
        if (!$this->trustHtml) {
            $html = $this->purifyHTML(
                $html,
                [
                    'HTML.Doctype' => 'HTML 4.01 Transitional',
                    'HTML.Trusted' => true,
                    'CSS.Trusted'  => true,
                ]
            );
        }

        $src_doc = new \DOMDocument();
        $src_doc->loadHTML('<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
</head><body>' . $html . "</body></html>");
        $bodies = $src_doc->getElementsByTagName('body');

        return $bodies->item(0);
    }

    public function setDebug(bool $debug): void
    {
        $this->DEBUG = $debug;
    }

    public function debugOutput(): void
    {
        $this->doc->preserveWhiteSpace = false;
        $this->doc->formatOutput       = true;
        echo htmlentities($this->doc->saveXML(), ENT_COMPAT, 'UTF-8');
        die();
    }

    /**
     * @param string[] $attributes
     */
    protected function appendStyleNode(string $styleName, string $family, string $element, array $attributes): void
    {
        $node = $this->doc->createElementNS(static::NS_STYLE, 'style');
        $node->setAttribute('style:name', $styleName);
        $node->setAttribute('style:family', $family);

        $style = $this->doc->createElementNS(static::NS_STYLE, $element);
        foreach ($attributes as $att_name => $att_val) {
            $style->setAttribute($att_name, $att_val);
        }
        $node->appendChild($style);

        foreach ($this->doc->getElementsByTagNameNS(static::NS_OFFICE, 'automatic-styles') as $element) {
            /** @var \DOMElement $element */
            $element->appendChild($node);
        }
    }

    protected function appendTextStyleNode(string $styleName, array $attributes): void
    {
        $this->appendStyleNode($styleName, 'text', 'text-properties', $attributes);
    }

    protected function appendParagraphStyleNode(string $styleName, array $attributes): void
    {
        $this->appendStyleNode($styleName, 'paragraph', 'paragraph-properties', $attributes);
    }

    /**
     * @param string $top (e.g. "20mm")
     */
    public function setMargins(string $top, string $left, string $right, string $bottom): void
    {
        $this->marginBottom = $bottom;
        $this->marginLeft   = $left;
        $this->marginRight  = $right;
        $this->marginTop    = $top;
    }

    /**
     * @param string $width (e.g. "297mm")
     * @param string $height (e.g. "210mm")
     * @param string $orientation (landscape or portrait)
     */
    public function setPageOrientation(string $width, string $height, string $orientation): void
    {
        $this->pageHeight       = $height;
        $this->pageWidth        = $width;
        $this->printOrientation = $orientation;
    }
}
