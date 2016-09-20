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

    /**
     * @param array $options
     *
     * @throws \Exception
     */
    public function __construct($options = [])
    {
        if (isset($options['templateFile']) && $options['templateFile'] != '') {
            $templateFile = $options['templateFile'];
        } else {
            $templateFile = __DIR__ . DIRECTORY_SEPARATOR . 'default-template.odt';
        }
        parent::__construct($templateFile, $options);
    }

    /**
     * @param string $filename
     */
    public function finishAndOutputOdt($filename = '')
    {
        header('Content-Type: application/vnd.oasis.opendocument.text');
        if ($filename != '') {
            header('Content-disposition: attachment;filename="' . addslashes($filename) . '"');
        }

        echo $this->finishAndGetDocument();

        die();
    }

    /**
     * @param string $search
     * @param string $replace
     */
    public function addReplace($search, $replace)
    {
        $this->replaces[$search] = $replace;
    }

    /**
     * @param string $html
     * @param bool $lineNumbered
     */
    public function addHtmlTextBlock($html, $lineNumbered = false)
    {
        $this->textBlocks[] = ['text' => $html, 'lineNumbered' => $lineNumbered];
    }

    /**
     * @param \DOMNode $srcNode
     * @param bool $lineNumbered
     * @param bool $inP
     *
     * @return \DOMNode[]
     * @throws \Exception
     */
    protected function html2ooNodeInt($srcNode, $lineNumbered, $inP)
    {
        $retNodes = [];
        switch ($srcNode->nodeType) {
            case XML_ELEMENT_NODE:
                /** @var \DOMElement $srcNode */
                if ($this->DEBUG) {
                    echo "Element - " . $srcNode->nodeName . " / Children: " . count($srcNode->childNodes) . "<br>";
                }
                $needsIntermediateP = false;
                switch ($srcNode->nodeName) {
                    case 'span':
                        $dstEl = $this->doc->createElementNS(static::NS_TEXT, 'span');
                        if ($srcNode->hasAttribute('class')) {
                            $classes = explode(' ', $srcNode->getAttribute('class'));
                            if (in_array('underline', $classes)) {
                                $dstEl->setAttribute('text:style-name', 'AntragsgruenUnderlined');
                            }
                            if (in_array('strike', $classes)) {
                                $dstEl->setAttribute('text:style-name', 'AntragsgruenStrike');
                            }
                            if (in_array('ins', $classes)) {
                                $dstEl->setAttribute('text:style-name', 'AntragsgruenIns');
                            }
                            if (in_array('inserted', $classes)) {
                                $dstEl->setAttribute('text:style-name', 'AntragsgruenIns');
                            }
                            if (in_array('del', $classes)) {
                                $dstEl->setAttribute('text:style-name', 'AntragsgruenDel');
                            }
                            if (in_array('deleted', $classes)) {
                                $dstEl->setAttribute('text:style-name', 'AntragsgruenDel');
                            }
                            if (in_array('superscript', $classes)) {
                                $dstEl->setAttribute('text:style-name', 'AntragsgruenSup');
                            }
                            if (in_array('subscritp', $classes)) {
                                $dstEl->setAttribute('text:style-name', 'AntragsgruenSub');
                            }
                        }
                        break;
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
                    case 'div':
                        if ($inP) {
                            $dstEl = $this->createNodeWithBaseStyle('span', $lineNumbered);
                        } else {
                            $dstEl = $this->createNodeWithBaseStyle('p', $lineNumbered);
                        }
                        $inP   = true;
                        break;
                    case 'blockquote':
                        $dstEl = $this->createNodeWithBaseStyle('p', $lineNumbered);
                        $class = ($lineNumbered ? 'Blockquote_Linenumbered' : 'Blockquote');
                        $dstEl->setAttribute('text:style-name', 'Antragsgrün_20_' . $class);
                        if (count($srcNode->childNodes) == 1) {
                            foreach ($srcNode->childNodes as $child) {
                                if ($child->nodeName == 'p') {
                                    $srcNode = $child;
                                }
                            }
                        }
                        $inP = true;
                        break;
                    case 'ul':
                        $dstEl = $this->doc->createElementNS(static::NS_TEXT, 'list');
                        break;
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
                    default:
                        throw new \Exception('Unknown Tag: ' . $srcNode->nodeName);
                }
                foreach ($srcNode->childNodes as $child) {
                    /** @var \DOMNode $child */
                    if ($this->DEBUG) {
                        echo "CHILD<br>" . $child->nodeType . "<br>";
                    }

                    $dstNodes = $this->html2ooNodeInt($child, $lineNumbered, $inP);
                    if ($this->DEBUG) {
                        echo "CHILD";
                        var_dump($dstNodes);
                    }
                    if ($dstNodes) {
                        foreach ($dstNodes as $dstNode) {
                            if ($needsIntermediateP && !in_array(strtolower($child->nodeName), ['p', 'ul', 'ol'])) {
                                $appendNode = $this->getNextNodeTemplate($lineNumbered);
                                $appendNode->appendChild($dstNode);
                                $dstEl->appendChild($appendNode);
                            } else {
                                $dstEl->appendChild($dstNode);
                            }
                        }
                    }
                }
                $retNodes[] = $dstEl;
                break;
            case XML_TEXT_NODE:
                /** @var \DOMText $srcNode */
                $textnode       = new \DOMText();
                $textnode->data = $srcNode->data;
                if ($this->DEBUG) {
                    echo 'Text<br>';
                }
                $retNodes[] = $textnode;
                break;
            case XML_DOCUMENT_TYPE_NODE:
                if ($this->DEBUG) {
                    echo 'Type Node<br>';
                }
                break;
            default:
                if ($this->DEBUG) {
                    echo 'Unknown Node: ' . $srcNode->nodeType . '<br>';
                }
        }

        return $retNodes;
    }

    /**
     * @param string $html
     * @param bool $lineNumbered
     *
     * @return \DOMNode[]
     */
    protected function html2ooNodes($html, $lineNumbered)
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

    /**
     * @return string
     */
    public function create()
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
            'fo:color'                   => '#00ff00',
            'style:text-underline-style' => 'solid',
            'style:text-underline-width' => 'auto',
            'style:text-underline-color' => 'font-color',
            'fo:font-weight'             => 'bold',
            'style:font-weight-asian'    => 'bold',
            'style:font-weight-complex'  => 'bold',
        ]);
        $this->appendTextStyleNode('AntragsgruenDel', [
            'fo:color'                      => '#ff0000',
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

        return $this->doc->saveXML();
    }

    /**
     * @param bool $lineNumbers
     *
     * @return \DOMNode
     */
    protected function getNextNodeTemplate($lineNumbers)
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
     * @param string $nodeType
     * @param bool $lineNumbers
     *
     * @return \DOMElement|\DOMNode
     */
    protected function createNodeWithBaseStyle($nodeType, $lineNumbers)
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
