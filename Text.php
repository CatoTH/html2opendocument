<?php

/**
 * @link https://github.com/CatoTH/html2opendocument
 * @author Tobias Hößl <tobias@hoessl.eu>
 * @license https://opensource.org/licenses/MIT
 */

namespace CatoTH\HTML2OpenDocument;

class Text extends Base
{
    /** @var null|\DOMElement */
    private $nodeText = null;

    /** @var bool */
    private $node_template_1_used = false;

    /** @var string[] */
    private $replaces = [];

    /** @var array */
    private $textBlocks = [];

    /** @var null|\Closure */
    protected $preSaveHook = null;

    const STYLE_INS = 'ins';
    const STYLE_DEL = 'del';

    /**
     * @throws \Exception
     */
    public function __construct(array $options = [])
    {
        if (isset($options['templateFile']) && $options['templateFile'] != '') {
            $templateFile = $options['templateFile'];
        } else {
            $templateFile = __DIR__ . DIRECTORY_SEPARATOR . 'default-template.odt';
        }
        parent::__construct($templateFile, $options);
    }

    public function finishAndOutputOdt(string $filename = ''): void
    {
        header('Content-Type: application/vnd.oasis.opendocument.text');
        if ($filename != '') {
            header('Content-disposition: attachment;filename="' . addslashes($filename) . '"');
        }

        echo $this->finishAndGetDocument();

        die();
    }

    public function addReplace(string $search, string $replace): void
    {
        $this->replaces[$search] = $replace;
    }

    public function addHtmlTextBlock(string $html, bool $lineNumbered = false)
    {
        $this->textBlocks[] = ['text' => $html, 'lineNumbered' => $lineNumbered];
    }

    /**
     * @return string[]
     */
    protected static function getCSSClasses(\DOMElement $element): array
    {
        if ($element->hasAttribute('class')) {
            return explode(' ', $element->getAttribute('class'));
        } else {
            return [];
        }
    }

    /**
     * @param string[] $parentStyles
     * @return string[]
     */
    protected static function getChildStyles(\DOMElement $element, array $parentStyles = []): array
    {
        $classes     = static::getCSSClasses($element);
        $childStyles = $parentStyles;
        if (in_array('ins', $classes)) {
            $childStyles[] = static::STYLE_INS;
        }
        if (in_array('inserted', $classes)) {
            $childStyles[] = static::STYLE_INS;
        }
        if (in_array('del', $classes)) {
            $childStyles[] = static::STYLE_DEL;
        }
        if (in_array('deleted', $classes)) {
            $childStyles[] = static::STYLE_DEL;
        }
        return array_unique($childStyles);
    }

	/**
	 * @param string[] $classes
	 */
    protected static function cssClassesToInternalClass(array $classes): ?string
    {
        if (in_array('underline', $classes)) {
            return 'AntragsgruenUnderlined';
        }
        if (in_array('strike', $classes)) {
            return 'AntragsgruenStrike';
        }
        if (in_array('ins', $classes)) {
            return 'AntragsgruenIns';
        }
        if (in_array('inserted', $classes)) {
            return 'AntragsgruenIns';
        }
        if (in_array('del', $classes)) {
            return 'AntragsgruenDel';
        }
        if (in_array('deleted', $classes)) {
            return 'AntragsgruenDel';
        }
        if (in_array('superscript', $classes)) {
            return 'AntragsgruenSup';
        }
        if (in_array('subscript', $classes)) {
            return 'AntragsgruenSub';
        }
        return null;
    }

    /**
     * Wraps all child nodes with text:p nodes, if necessary
     * (it's not necessary for child nodes that are p's themselves or lists)
     */
    protected function wrapChildrenWithP(\DOMElement $parentEl, bool $lineNumbered): \DOMElement
    {
        $childNodes = [];
        while ($parentEl->childNodes->length > 0) {
            $el = $parentEl->firstChild;
            $parentEl->removeChild($el);
            $childNodes[] = $el;
        }

        $appendNode = null;
        foreach ($childNodes as $childNode) {
            if (in_array(strtolower($childNode->nodeName), ['p', 'list'])) {
                if ($appendNode) {
                    $parentEl->appendChild($appendNode);
                    $appendNode = null;
                }
                $parentEl->appendChild($childNode);
            } else {
                if (!$appendNode) {
                    $appendNode = $this->getNextNodeTemplate($lineNumbered);
                }
                $appendNode->appendChild($childNode);
            }
        }
        if ($appendNode) {
            $parentEl->appendChild($appendNode);
        }

        return $parentEl;
    }

    /**
     * @param \DOMNode $srcNode
     * @param bool $lineNumbered
     * @param bool $inP
     * @param string[] $parentStyles
     *
     * @return \DOMNode[]
     * @throws \Exception
     */
    protected function html2ooNodeInt(\DOMNode $srcNode, bool $lineNumbered, bool $inP, array $parentStyles = []): array
    {
        switch ($srcNode->nodeType) {
            case XML_ELEMENT_NODE:
                /** @var \DOMElement $srcNode */
                if ($this->DEBUG) {
                    echo "Element - " . $srcNode->nodeName . " / Children: " . $srcNode->childNodes->length . "<br>";
                }
                $needsIntermediateP = false;
                $childStyles        = static::getChildStyles($srcNode, $parentStyles);
                switch ($srcNode->nodeName) {
                    case 'b':
                    case 'strong':
                        $dstEl = $this->doc->createElementNS(static::NS_TEXT, 'span');
                        $dstEl->setAttribute('text:style-name', 'AntragsgruenBold');
                        break;
                    case 'i':
                    case 'em':
                        $dstEl = $this->doc->createElementNS(static::NS_TEXT, 'span');
                        $dstEl->setAttribute('text:style-name', 'AntragsgruenItalic');
                        break;
                    case 's':
                        $dstEl = $this->doc->createElementNS(static::NS_TEXT, 'span');
                        $dstEl->setAttribute('text:style-name', 'AntragsgruenStrike');
                        break;
                    case 'u':
                        $dstEl = $this->doc->createElementNS(static::NS_TEXT, 'span');
                        $dstEl->setAttribute('text:style-name', 'AntragsgruenUnderlined');
                        break;
                    case 'sub':
                        $dstEl = $this->doc->createElementNS(static::NS_TEXT, 'span');
                        $dstEl->setAttribute('text:style-name', 'AntragsgruenSub');
                        break;
                    case 'sup':
                        $dstEl = $this->doc->createElementNS(static::NS_TEXT, 'span');
                        $dstEl->setAttribute('text:style-name', 'AntragsgruenSup');
                        break;
                    case 'br':
                        $dstEl = $this->doc->createElementNS(static::NS_TEXT, 'line-break');
                        break;
                    case 'del':
                        $dstEl = $this->doc->createElementNS(static::NS_TEXT, 'span');
                        $dstEl->setAttribute('text:style-name', 'AntragsgruenDel');
                        break;
                    case 'ins':
                        $dstEl = $this->doc->createElementNS(static::NS_TEXT, 'span');
                        $dstEl->setAttribute('text:style-name', 'AntragsgruenIns');
                        break;
                    case 'a':
                        $dstEl = $this->doc->createElementNS(static::NS_TEXT, 'a');
                        try {
                            $attr = $srcNode->getAttribute('href');
                            if ($attr) {
                                $dstEl->setAttribute('xlink:href', $attr);
                            }
                        } catch (\Exception $e) {
                        }
                        break;
                    case 'p':
                        if ($inP) {
                            $dstEl = $this->createNodeWithBaseStyle('span', $lineNumbered);
                        } else {
                            $dstEl = $this->createNodeWithBaseStyle('p', $lineNumbered);
                        }
                        $intClass = static::cssClassesToInternalClass(static::getCSSClasses($srcNode));
                        if ($intClass) {
                            $dstEl->setAttribute('text:style-name', $intClass);
                        }
                        $inP = true;
                        break;
                    case 'div':
                        // We're basically ignoring DIVs here, as there is no corresponding element in OpenDocument
                        // Therefore no support for styles and classes set on DIVs yet.
                        $dstEl = null;
                        break;
                    case 'blockquote':
                        $dstEl = $this->createNodeWithBaseStyle('p', $lineNumbered);
                        $class = ($lineNumbered ? 'Blockquote_Linenumbered' : 'Blockquote');
                        $dstEl->setAttribute('text:style-name', 'Antragsgrün_20_' . $class);
                        if ($srcNode->childNodes->length == 1) {
                            foreach ($srcNode->childNodes as $child) {
                                if ($child->nodeName == 'p') {
                                    $srcNode = $child;
                                }
                            }
                        }
                        $inP = true;
                        break;
                    case 'ul':
                    case 'ol':
                        $dstEl = $this->doc->createElementNS(static::NS_TEXT, 'list');
                        break;
                    case 'li':
                        $dstEl              = $this->doc->createElementNS(static::NS_TEXT, 'list-item');
                        $needsIntermediateP = true;
                        $inP                = true;
                        break;
                    case 'h1':
                        $dstEl = $this->createNodeWithBaseStyle('p', $lineNumbered);
                        $dstEl->setAttribute('text:style-name', 'Antragsgrün_20_H1');
                        $inP = true;
                        break;
                    case 'h2':
                        $dstEl = $this->createNodeWithBaseStyle('p', $lineNumbered);
                        $dstEl->setAttribute('text:style-name', 'Antragsgrün_20_H2');
                        $inP = true;
                        break;
                    case 'h3':
                        $dstEl = $this->createNodeWithBaseStyle('p', $lineNumbered);
                        $dstEl->setAttribute('text:style-name', 'Antragsgrün_20_H3');
                        $inP = true;
                        break;
                    case 'h4':
                    case 'h5':
                    case 'h6':
                        $dstEl = $this->createNodeWithBaseStyle('p', $lineNumbered);
                        $dstEl->setAttribute('text:style-name', 'Antragsgrün_20_H4');
                        $inP = true;
                        break;
                    case 'span':
                    default:
                        $dstEl    = $this->doc->createElementNS(static::NS_TEXT, 'span');
                        $intClass = static::cssClassesToInternalClass(static::getCSSClasses($srcNode));
                        if ($intClass) {
                            $dstEl->setAttribute('text:style-name', $intClass);
                        }
                        break;
                }


                if ($dstEl === null) {
                    $ret = [];
                    foreach ($srcNode->childNodes as $child) {
                        /** @var \DOMNode $child */
                        if ($this->DEBUG) {
                            echo "CHILD<br>" . $child->nodeType . "<br>";
                        }

                        $dstNodes = $this->html2ooNodeInt($child, $lineNumbered, $inP, $childStyles);
                        foreach ($dstNodes as $dstNode) {
                            $ret[] = $dstNode;
                        }
                    }
                    return $ret;
                }

                foreach ($srcNode->childNodes as $child) {
                    /** @var \DOMNode $child */
                    if ($this->DEBUG) {
                        echo "CHILD<br>" . $child->nodeType . "<br>";
                    }

                    $dstNodes = $this->html2ooNodeInt($child, $lineNumbered, $inP, $childStyles);
                    foreach ($dstNodes as $dstNode) {
                        $dstEl->appendChild($dstNode);
                    }
                }

                if ($needsIntermediateP && $dstEl->childNodes->length > 0) {
                    $dstEl = static::wrapChildrenWithP($dstEl, $lineNumbered);
                }
                return [$dstEl];
            case XML_TEXT_NODE:
                /** @var \DOMText $srcNode */
                $textnode       = new \DOMText();
                $textnode->data = $srcNode->data;
                if ($this->DEBUG) {
                    echo 'Text<br>';
                }
                if (in_array(static::STYLE_DEL, $parentStyles)) {
                    $dstEl = $this->createNodeWithBaseStyle('span', $lineNumbered);
                    $dstEl->setAttribute('text:style-name', 'AntragsgruenDel');
                    $dstEl->appendChild($textnode);
                    $textnode = $dstEl;
                }
                if (in_array(static::STYLE_INS, $parentStyles)) {
                    $dstEl = $this->createNodeWithBaseStyle('span', $lineNumbered);
                    $dstEl->setAttribute('text:style-name', 'AntragsgruenIns');
                    $dstEl->appendChild($textnode);
                    $textnode = $dstEl;
                }
                return [$textnode];
                break;
            case XML_DOCUMENT_TYPE_NODE:
                if ($this->DEBUG) {
                    echo 'Type Node<br>';
                }
                return [];
            default:
                if ($this->DEBUG) {
                    echo 'Unknown Node: ' . $srcNode->nodeType . '<br>';
                }
                return [];
        }
    }

	/**
	 * @return \DOMNode[]
	 * @throws \Exception
	 */
    protected function html2ooNodes(string $html, bool $lineNumbered): array
    {
        if (!is_string($html)) {
            echo print_r($html, true);
            echo print_r(debug_backtrace(DEBUG_BACKTRACE_IGNORE_ARGS), true);
            die();
        }

        $body = $this->html2DOM($html);

        $retNodes = [];
        for ($i = 0; $i < $body->childNodes->length; $i++) {
            $child = $body->childNodes->item($i);

            /** @var \DOMNode $child */
            if ($child->nodeName == 'ul') {
                // Alle anderen Nodes dieses Aufrufs werden ignoriert
                if ($this->DEBUG) {
                    echo 'LIST<br>';
                }
                $recNewNodes = $this->html2ooNodeInt($child, $lineNumbered, false);
            } else {
                if ($child->nodeType == XML_TEXT_NODE) {
                    $new_node = $this->getNextNodeTemplate($lineNumbered);
                    /** @var \DOMText $child */
                    if ($this->DEBUG) {
                        echo $child->nodeName . ' - ' . htmlentities($child->data, ENT_COMPAT, 'UTF-8') . '<br>';
                    }
                    $text       = new \DOMText();
                    $text->data = $child->data;
                    $new_node->appendChild($text);
                    $recNewNodes = [$new_node];
                } else {
                    if ($this->DEBUG) {
                        echo $child->nodeName . '!!!!!!!!!!!!<br>';
                    }
                    $recNewNodes = $this->html2ooNodeInt($child, $lineNumbered, false);
                }
            }
            foreach ($recNewNodes as $recNewNode) {
                $retNodes[] = $recNewNode;
            }
        }

        return $retNodes;
    }

    public function setPreSaveHook(callable $cb): void {
        $this->preSaveHook = $cb;
    }

	/**
	 * @throws \Exception
	 */
    public function create(): string
    {
        $this->appendTextStyleNode('AntragsgruenBold', [
            'fo:font-weight'            => 'bold',
            'style:font-weight-asian'   => 'bold',
            'style:font-weight-complex' => 'bold',
        ]);
        $this->appendTextStyleNode('AntragsgruenItalic', [
            'fo:font-style'            => 'italic',
            'style:font-style-asian'   => 'italic',
            'style:font-style-complex' => 'italic',
        ]);
        $this->appendTextStyleNode('AntragsgruenUnderlined', [
            'style:text-underline-width' => 'auto',
            'style:text-underline-color' => 'font-color',
            'style:text-underline-style' => 'solid',
        ]);
        $this->appendTextStyleNode('AntragsgruenStrike', [
            'style:text-line-through-style' => 'solid',
            'style:text-line-through-type'  => 'single',
        ]);
        $this->appendTextStyleNode('AntragsgruenIns', [
            'fo:color'                   => $this->colorIns,
            'style:text-underline-style' => 'solid',
            'style:text-underline-width' => 'auto',
            'style:text-underline-color' => 'font-color',
            'fo:font-weight'             => 'bold',
            'style:font-weight-asian'    => 'bold',
            'style:font-weight-complex'  => 'bold',
        ]);
        $this->appendTextStyleNode('AntragsgruenDel', [
            'fo:color'                      => $this->colorDel,
            'style:text-line-through-style' => 'solid',
            'style:text-line-through-type'  => 'single',
            'fo:font-style'                 => 'italic',
            'style:font-style-asian'        => 'italic',
            'style:font-style-complex'      => 'italic',
        ]);
        $this->appendTextStyleNode('AntragsgruenSub', [
            'style:text-position' => 'sub 58%',
        ]);
        $this->appendTextStyleNode('AntragsgruenSup', [
            'style:text-position' => 'super 58%',
        ]);

        /** @var \DOMNode[] $nodes */
        $nodes = [];
        foreach ($this->doc->getElementsByTagNameNS(static::NS_TEXT, 'span') as $element) {
            $nodes[] = $element;
        }
        foreach ($this->doc->getElementsByTagNameNS(static::NS_TEXT, 'p') as $element) {
            $nodes[] = $element;
        }


        $searchFor   = array_keys($this->replaces);
        $replaceWith = array_values($this->replaces);
        foreach ($nodes as $node) {
            $children = $node->childNodes;
            foreach ($children as $child) {
                if ($child->nodeType == XML_TEXT_NODE) {
                    /** @var \DOMText $child */
                    $child->data = preg_replace($searchFor, $replaceWith, $child->data);

                    if (preg_match("/\{\{ANTRAGSGRUEN:DUMMY\}\}/siu", $child->data)) {
                        $node->parentNode->removeChild($node);
                    }
                    if (preg_match("/\{\{ANTRAGSGRUEN:TEXT\}\}/siu", $child->data)) {
                        $this->nodeText = $node;
                    }
                }
            }
        }

        foreach ($this->textBlocks as $textBlock) {
            $newNodes = $this->html2ooNodes($textBlock['text'], $textBlock['lineNumbered']);
            foreach ($newNodes as $newNode) {
                $this->nodeText->parentNode->insertBefore($newNode, $this->nodeText);
            }
        }

        $this->nodeText->parentNode->removeChild($this->nodeText);

        if ($this->preSaveHook !== null) {
            call_user_func($this->preSaveHook, $this->doc);
        }

        return $this->doc->saveXML();
    }

    protected function getNextNodeTemplate(bool $lineNumbers): \DOMNode
    {
        $node = $this->nodeText->cloneNode();
        /** @var \DOMElement $node */
        if ($lineNumbers) {
            if ($this->node_template_1_used) {
                $node->setAttribute('text:style-name', 'Antragsgrün_20_LineNumbered_20_Standard');
            } else {
                $this->node_template_1_used = true;
                $node->setAttribute('text:style-name', 'Antragsgrün_20_LineNumbered_20_First');
            }
        } else {
            $node->setAttribute('text:style-name', 'Antragsgrün_20_Standard');
        }

        return $node;
    }

    /**
     * @return \DOMElement|\DOMNode
     */
    protected function createNodeWithBaseStyle(string $nodeType, bool $lineNumbers)
    {
        $node = $this->doc->createElementNS(static::NS_TEXT, $nodeType);
        if ($lineNumbers) {
            if ($this->node_template_1_used) {
                $node->setAttribute('text:style-name', 'Antragsgrün_20_LineNumbered_20_Standard');
            } else {
                $this->node_template_1_used = true;
                $node->setAttribute('text:style-name', 'Antragsgrün_20_LineNumbered_20_First');
            }
        } else {
            $node->setAttribute('text:style-name', 'Antragsgrün_20_Standard');
        }

        return $node;
    }
}
