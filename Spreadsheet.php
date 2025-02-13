<?php

/**
 * @link https://github.com/CatoTH/html2opendocument
 * @author Tobias Hößl <tobias@hoessl.eu>
 * @license https://opensource.org/licenses/MIT
 */

namespace CatoTH\HTML2OpenDocument;

class Spreadsheet extends Base
{
    public const TYPE_TEXT   = 0;
    public const TYPE_NUMBER = 1;
    public const TYPE_HTML   = 2;
    public const TYPE_LINK   = 3;
    
    public const FORMAT_LINEBREAK  = 0;
    public const FORMAT_BOLD       = 1;
    public const FORMAT_ITALIC     = 2;
    public const FORMAT_UNDERLINED = 3;
    public const FORMAT_STRIKE     = 4;
    public const FORMAT_INS        = 5;
    public const FORMAT_DEL        = 6;
    public const FORMAT_LINK       = 7;
    public const FORMAT_INDENTED   = 8;
    public const FORMAT_SUP        = 9;
    public const FORMAT_SUB        = 10;

    public static array $FORMAT_NAMES = [
        0  => 'linebreak',
        1  => 'bold',
        2  => 'italic',
        3  => 'underlined',
        4  => 'strike',
        5  => 'ins',
        6  => 'del',
        7  => 'link',
        8  => 'indented',
        9  => 'sup',
        10 => 'sub',
    ];
    
    protected ?\DOMDocument $doc = NULL;
    
    protected \DOMElement $domTable;
    
    protected array $matrix           = [];
    protected int   $matrixRows       = 0;
    protected int   $matrixCols       = 0;
    protected array $matrixColWidths  = [];
    protected array $matrixRowHeights = [];
    
    protected array $rowNodes         = [];
    protected array $cellNodeMatrix   = [];
    protected array $cellStylesMatrix = [];
    
    protected array $classCache = [];
    
    protected ?\Closure $preSaveHook = NULL;

    /**
     * @param array $options
     * @throws \Exception
     */
    public function __construct($options = [])
    {
        if (isset($options['templateFile']) && is_string($options['templateFile']) && $options['templateFile'] !== '') {
            $templateFile = $options['templateFile'];
        } else {
            $templateFile = __DIR__ . DIRECTORY_SEPARATOR . 'default-template.ods';
        }
        parent::__construct($templateFile, $options);
    }

    public function finishAndOutputOds(string $filename = ''): void
    {
        header('Content-Type: application/vnd.oasis.opendocument.spreadsheet');
        if ($filename !== '') {
            header('Content-disposition: attachment;filename="' . addslashes($filename) . '"');
        }

        echo $this->finishAndGetDocument();

        die();
    }

    protected function appendCellStyleNode(string $styleName, array $cellAttributes, array $textAttributes): void
    {
        $node = $this->doc->createElementNS(static::NS_STYLE, "style");
        $node->setAttribute("style:name", $styleName);
        $node->setAttribute("style:family", 'table-cell');
        $node->setAttribute("style:parent-style-name", "Default");

        if (count($cellAttributes) > 0) {
            $style = $this->doc->createElementNS(static::NS_STYLE, 'table-cell-properties');
            foreach ($cellAttributes as $att_name => $att_val) {
                $style->setAttribute($att_name, $att_val);
            }
            $node->appendChild($style);
        }
        if (count($textAttributes) > 0) {
            $style = $this->doc->createElementNS(static::NS_STYLE, 'text-properties');
            foreach ($textAttributes as $att_name => $att_val) {
                $style->setAttribute($att_name, $att_val);
            }
            $node->appendChild($style);
        }

        foreach ($this->doc->getElementsByTagNameNS(static::NS_OFFICE, 'automatic-styles') as $element) {
            /** @var \DOMElement $element */
            $element->appendChild($node);
        }
    }

    protected function appendColStyleNode(string $styleName, array $attributes): void
    {
        $this->appendStyleNode($styleName, 'table-column', 'table-column-properties', $attributes);
    }

    protected function appendRowStyleNode(string $styleName, array $attributes): void
    {
        $this->appendStyleNode($styleName, 'table-row', 'table-row-properties', $attributes);
    }

    protected function initRow(int $row): void
    {
        if (!isset($this->matrix[$row])) {
            $this->matrix[$row] = [];
        }
        if ($row > $this->matrixRows) {
            $this->matrixRows = $row;
        }
    }

    /**
     * @param mixed $content:
     * - for contentType === Spreadsheet::TYPE_LINK: ['href' => $href, 'text' => $email]
     * - for all other types: a string
     */
    public function setCell(int $row, int $col, int $contentType, mixed $content, ?string $cssClass = null, ?array $styles = null): void
    {
        $this->initRow($row);
        if ($col > $this->matrixCols) {
            $this->matrixCols = $col;
        }
        $this->matrix[$row][$col] = [
            'type'    => $contentType,
            'content' => $content,
            'class'   => $cssClass,
            'styles'  => $styles,
        ];
    }

    public function setColumnWidth(int $col, float $widthInCm): void
    {
        $this->matrixColWidths[$col] = $widthInCm;
    }

    public function setMinRowHeight(int $row, float $minHeightInCm): void
    {
        $this->initRow($row);
        $rowHeight = $this->matrixRowHeights[$row] ?? 1;
        if ($minHeightInCm > $rowHeight) {
            $rowHeight = $minHeightInCm;
        }
        $this->matrixRowHeights[$row] = $rowHeight;
    }

    /**
     * @throws \Exception
     */
    protected function getCleanDomTable(): \DOMElement
    {
        $domTables = $this->doc->getElementsByTagNameNS(static::NS_TABLE, 'table');
        if ($domTables->length !== 1) {
            throw new \RuntimeException('Could not parse ODS template.');
        }

        $this->domTable = $domTables->item(0);

        $children = $this->domTable->childNodes;
        for ($i = $children->length - 1; $i >= 0; $i--) {
            $this->domTable->removeChild($children->item($i));
        }
        return $this->domTable;
    }


    protected function setColStyles(): void
    {
        for ($col = 0; $col <= $this->matrixCols; $col++) {
            $element = $this->doc->createElementNS(static::NS_TABLE, 'table-column');
            if (isset($this->matrixColWidths[$col])) {
                $element->setAttribute('table:style-name', 'Antragsgruen_col_' . $col);
                $this->appendColStyleNode('Antragsgruen_col_' . $col, [
                    'style:column-width' => $this->matrixColWidths[$col] . 'cm',
                ]);
            }
            $this->domTable->appendChild($element);
        }
    }

    protected function setCellContent(): void
    {
        for ($row = 0; $row <= $this->matrixRows; $row++) {
            $this->cellNodeMatrix[$row] = [];
            $currentRow                 = $this->doc->createElementNS(static::NS_TABLE, 'table-row');
            for ($col = 0; $col <= $this->matrixCols; $col++) {
                $this->cellNodeMatrix[$row][$col] = [];
                $currentCell                      = $this->doc->createElementNS(static::NS_TABLE, 'table-cell');
                if (isset($this->matrix[$row][$col])) {
                    $cell = $this->matrix[$row][$col];
                    switch ($cell["type"]) {
                        case static::TYPE_TEXT:
                            $elementP              = $this->doc->createElementNS(static::NS_TEXT, 'p');
                            $elementP->textContent = $cell['content'];
                            $currentCell->appendChild($elementP);
                            break;
                        case static::TYPE_NUMBER:
                            $elementP              = $this->doc->createElementNS(static::NS_TEXT, 'p');
                            $elementP->textContent = $cell['content'];
                            $currentCell->appendChild($elementP);
                            $currentCell->setAttribute('calcext:value-type', 'float');
                            $currentCell->setAttribute('office:value-type', 'float');
                            $currentCell->setAttribute('office:value', (string)$cell['content']);
                            break;
                        case static::TYPE_LINK:
                            $elementP = $this->doc->createElementNS(static::NS_TEXT, 'p');
                            $elementA = $this->doc->createElementNS(static::NS_TEXT, 'a');
                            $elementA->setAttributeNS(static::NS_XLINK, 'xlink:href', $cell['content']['href']);
                            $textNode = $this->doc->createTextNode($cell['content']['text']);
                            $elementA->appendChild($textNode);
                            $elementP->appendChild($elementA);
                            $currentCell->appendChild($elementP);
                            break;
                        case static::TYPE_HTML:
                            $nodes = $this->html2OdsNodes($cell['content']);
                            foreach ($nodes as $node) {
                                $currentCell->appendChild($node);
                            }

                            //$this->setMinRowHeight($row, count($ps));
                            $styles = $cell['styles'];
                            if (isset($styles['fo:wrap-option']) && $styles['fo:wrap-option'] === 'no-wrap') {
                                $wrap   = 'no-wrap';
                                $height = 1;
                            } else {
                                $wrap   = 'wrap';
                                $width  = $this->matrixColWidths[$col] ?? 2;
                                $height = (mb_strlen(strip_tags($this->matrix[$row][$col]['content'])) / ($width * 6));
                            }
                            $this->setCellStyle($row, $col, [
                                'fo:wrap-option' => $wrap,
                            ], [
                                'fo:hyphenate' => 'true',
                            ]);
                            $this->setMinRowHeight($row, $height);
                            break;
                    }
                }
                $currentRow->appendChild($currentCell);
                $this->cellNodeMatrix[$row][$col] = $currentCell;
            }
            $this->domTable->appendChild($currentRow);
            $this->rowNodes[$row] = $currentRow;
        }
    }

    public function setCellStyle(int $row, int $col, ?array $cellAttributes, ?array $textAttributes): void
    {
        if (!isset($this->cellStylesMatrix[$row])) {
            $this->cellStylesMatrix[$row] = [];
        }
        if (!isset($this->cellStylesMatrix[$row][$col])) {
            $this->cellStylesMatrix[$row][$col] = ['cell' => [], 'text' => []];
        }
        if (is_array($cellAttributes)) {
            foreach ($cellAttributes as $key => $val) {
                $this->cellStylesMatrix[$row][$col]['cell'][$key] = $val;
            }
        }
        if (is_array($textAttributes)) {
            foreach ($textAttributes as $key => $val) {
                $this->cellStylesMatrix[$row][$col]['text'][$key] = $val;
            }
        }
    }

    public function setCellStyles(): void
    {
        for ($row = 0; $row <= $this->matrixRows; $row++) {
            for ($col = 0; $col <= $this->matrixCols; $col++) {
                if (isset($this->cellStylesMatrix[$row]) && isset($this->cellStylesMatrix[$row][$col])) {
                    $cell = $this->cellStylesMatrix[$row][$col];
                } else {
                    $cell = ['cell' => [], 'text' => []];
                }

                $styleId    = 'Antragsgruen_cell_' . $row . '_' . $col;
                $cellStyles = array_merge([
                    'style:vertical-align' => 'top'
                ], $cell['cell']);
                $this->appendCellStyleNode($styleId, $cellStyles, $cell['text']);
                /** @var \DOMElement $currentCell */
                $currentCell = $this->cellNodeMatrix[$row][$col];
                $currentCell->setAttribute('table:style-name', $styleId);
            }
        }
	    /*
        foreach ($this->cellStylesMatrix as $rowNr => $row) {
            foreach ($row as $colNr => $cell) {
            }
        }
	    */
    }

    public function setRowStyles(): void
    {
        foreach ($this->matrixRowHeights as $row => $height) {
            $styleName = 'Antragsgruen_row_' . $row;
            $this->appendRowStyleNode($styleName, [
                'style:row-height' => ($height * 0.45) . 'cm',
            ]);

            /** @var \DOMElement $node */
            $node = $this->rowNodes[$row];
            $node->setAttribute('table:style-name', $styleName);
        }
    }

    public function drawBorder(int $fromRow, int $fromCol, int $toRow, int $toCol, float $width): void
    {
        for ($i = $fromRow; $i <= $toRow; $i++) {
            $this->setCellStyle($i, $fromCol, [
                'fo:border-left' => $width . 'pt solid #000000',
            ], []);
            $this->setCellStyle($i, $toCol, [
                'fo:border-right' => $width . 'pt solid #000000',
            ], []);
        }

        for ($i = $fromCol; $i <= $toCol; $i++) {
            $this->setCellStyle($fromRow, $i, [
                'fo:border-top' => $width . 'pt solid #000000',
            ], []);
            $this->setCellStyle($toRow, $i, [
                'fo:border-bottom' => $width . 'pt solid #000000',
            ], []);
        }
    }

    protected function node2Formatting(\DOMElement $node, array $currentFormats): array
    {
        switch ($node->nodeName) {
            case 'b':
            case 'strong':
                $currentFormats[] = static::FORMAT_BOLD;
                break;
            case 'i':
            case 'em':
                $currentFormats[] = static::FORMAT_ITALIC;
                break;
            case 's':
                $currentFormats[] = static::FORMAT_STRIKE;
                break;
            case 'u':
                $currentFormats[] = static::FORMAT_UNDERLINED;
                break;
            case 'sub':
                $currentFormats[] = static::FORMAT_SUB;
                break;
            case 'sup':
                $currentFormats[] = static::FORMAT_SUP;
                break;
            case 'br':
                break;
            case 'p':
            case 'div':
            case 'blockquote':
                if ($node->hasAttribute('class')) {
                    $classes = explode(' ', $node->getAttribute('class'));
                    if (in_array('underline', $classes)) {
                        $currentFormats[] = static::FORMAT_UNDERLINED;
                    }
                    if (in_array('strike', $classes)) {
                        $currentFormats[] = static::FORMAT_STRIKE;
                    }
                    if (in_array('ins', $classes)) {
                        $currentFormats[] = static::FORMAT_INS;
                    }
                    if (in_array('inserted', $classes)) {
                        $currentFormats[] = static::FORMAT_INS;
                    }
                    if (in_array('del', $classes)) {
                        $currentFormats[] = static::FORMAT_DEL;
                    }
                    if (in_array('deleted', $classes)) {
                        $currentFormats[] = static::FORMAT_DEL;
                    }
                }
                break;
            case 'ul':
            case 'ol':
                if ($node->hasAttribute('class')) {
                    $classes          = explode(' ', $node->getAttribute('class'));
                    $currentFormats[] = static::FORMAT_INDENTED;
                    if (in_array('ins', $classes)) {
                        $currentFormats[] = static::FORMAT_INS;
                    }
                    if (in_array('inserted', $classes)) {
                        $currentFormats[] = static::FORMAT_INS;
                    }
                    if (in_array('del', $classes)) {
                        $currentFormats[] = static::FORMAT_DEL;
                    }
                    if (in_array('deleted', $classes)) {
                        $currentFormats[] = static::FORMAT_DEL;
                    }
                }
                break;
            case 'li':
                break;
            case 'del':
                $currentFormats[] = static::FORMAT_DEL;
                break;
            case 'ins':
                $currentFormats[] = static::FORMAT_INS;
                break;
            case 'h1':
            case 'h2':
            case 'h3':
            case 'h4':
            case 'h5':
            case 'h6':
                $currentFormats[] = static::FORMAT_BOLD;
                break;
            case 'a':
                $currentFormats[] = static::FORMAT_LINK;
                try {
                    $attr = $node->getAttribute('href');
                    if ($attr) {
                        $currentFormats['href'] = $attr;
                    }
                } catch (\Exception) {
                }
                break;
            case 'span':
            default:
                if ($node->hasAttribute('class')) {
                    $classes = explode(' ', $node->getAttribute('class'));
                    if (in_array('underline', $classes)) {
                        $currentFormats[] = static::FORMAT_UNDERLINED;
                    }
                    if (in_array('strike', $classes)) {
                        $currentFormats[] = static::FORMAT_STRIKE;
                    }
                    if (in_array('ins', $classes)) {
                        $currentFormats[] = static::FORMAT_INS;
                    }
                    if (in_array('inserted', $classes)) {
                        $currentFormats[] = static::FORMAT_INS;
                    }
                    if (in_array('del', $classes)) {
                        $currentFormats[] = static::FORMAT_DEL;
                    }
                    if (in_array('deleted', $classes)) {
                        $currentFormats[] = static::FORMAT_DEL;
                    }
                    if (in_array('superscript', $classes)) {
                        $currentFormats[] = static::FORMAT_SUP;
                    }
                    if (in_array('subscript', $classes)) {
                        $currentFormats[] = static::FORMAT_SUB;
                    }
                }
                break;
        }
        return $currentFormats;
    }

    protected function tokenizeFlattenHtml(\DOMNode $node, array $currentFormats): array
    {
        $return = [];
        foreach ($node->childNodes as $child) {
            switch ($child->nodeType) {
                case XML_ELEMENT_NODE:
                    /** @var \DOMElement $child */
                    $formattings = $this->node2Formatting($child, $currentFormats);
                    $children    = $this->tokenizeFlattenHtml($child, $formattings);
                    $return      = array_merge($return, $children);
                    if (in_array($child->nodeName, ['br', 'div', 'p', 'li', 'blockquote'])) {
                        $return[] = [
                            'text'        => '',
                            'formattings' => [static::FORMAT_LINEBREAK],
                        ];
                    }
                    if (in_array($child->nodeName, ['h1', 'h2', 'h3', 'h4', 'h5', 'h6'])) {
                        $return[] = [
                            'text'        => '',
                            'formattings' => [static::FORMAT_LINEBREAK, static::FORMAT_BOLD],
                        ];
                    }
                    break;
                case XML_TEXT_NODE:
                    /** @var \DOMText $child */
                    $return[] = [
                        'text'        => $child->data,
                        'formattings' => $currentFormats,
                    ];
                    break;
                default:
            }
        }

        return $return;
    }

    protected function getClassByFormats(array $formats): string
    {
        sort($formats);
        $key = implode('_', $formats);
        if (!isset($this->classCache[$key])) {
            $name   = 'Antragsgruen';
            $styles = [];
            foreach ($formats as $format) {
                if (!isset(static::$FORMAT_NAMES[$format])) {
                    continue;
                }
                $name .= '_' . static::$FORMAT_NAMES[$format];
                switch ($format) {
                    case static::FORMAT_INS:
                        $styles['fo:color']                   = $this->colorIns;
                        $styles['style:text-underline-style'] = 'solid';
                        $styles['style:text-underline-width'] = 'auto';
                        $styles['style:text-underline-color'] = 'font-color';
                        break;
                    case static::FORMAT_DEL:
                        $styles['fo:color']                     = $this->colorDel;
                        $styles['style:text-line-through-type'] = 'single';
                        break;
                    case static::FORMAT_BOLD:
                        $styles['fo:font-weight']            = 'bold';
                        $styles['style:font-weight-asian']   = 'bold';
                        $styles['style:font-weight-complex'] = 'bold';
                        break;
                    case static::FORMAT_UNDERLINED:
                        $styles['style:text-underline-width'] = 'auto';
                        $styles['style:text-underline-color'] = 'font-color';
                        $styles['style:text-underline-style'] = 'solid';
                        break;
                    case static::FORMAT_STRIKE:
                        $styles['style:text-line-through-type'] = 'single';
                        break;
                    case static::FORMAT_ITALIC:
                        $styles['fo:font-style']            = 'italic';
                        $styles['style:font-style-asian']   = 'italic';
                        $styles['style:font-style-complex'] = 'italic';
                        break;
                    case static::FORMAT_SUP:
                        $styles['fo:font-size']        = '10pt';
                        $styles['style:text-position'] = '31%';
                        break;
                    case static::FORMAT_SUB:
                        $styles['fo:font-size']        = '10pt';
                        $styles['style:text-position'] = '-31%';
                        break;
                }
            }
            $this->appendTextStyleNode($name, $styles);
            $this->classCache[$key] = $name;
        }
        return $this->classCache[$key];
    }

    public function html2OdsNodes(string $html): array
    {
        $body     = $this->html2DOM($html);
        $tokens   = $this->tokenizeFlattenHtml($body, []);
        $nodes    = [];
        $currentP = $this->doc->createElementNS(static::NS_TEXT, 'p');
        foreach ($tokens as $token) {
            if (trim($token['text']) !== '') {
                $node = $this->doc->createElement('text:span');
                if (count($token['formattings']) > 0) {
                    $className = $this->getClassByFormats($token['formattings']);
                    $node->setAttribute('text:style-name', $className);
                }
                $textNode = $this->doc->createTextNode($token['text']);
                $node->appendChild($textNode);
                $currentP->appendChild($node);
            }

            if (in_array(static::FORMAT_LINEBREAK, $token['formattings'])) {
                $nodes[]  = $currentP;
                $currentP = $this->doc->createElementNS(static::NS_TEXT, 'p');
            }
        }
        $nodes[] = $currentP;
        return $nodes;
    }

    public function setPreSaveHook(callable $cb): void
    {
        $this->preSaveHook = $cb;
    }


    /**
     * @throws \Exception
     */
    public function create(): string
    {
        $this->getCleanDomTable();
        $this->setColStyles();
        $this->setCellContent();
        $this->setRowStyles();
        $this->setCellStyles();

        if ($this->preSaveHook !== null) {
            call_user_func($this->preSaveHook, $this->doc);
        }

        $xml = $this->doc->saveXML();

        $rows = explode("\n", $xml);
        $rows[0] .= "\n";
        return implode('', $rows) . "\n";
    }
}
