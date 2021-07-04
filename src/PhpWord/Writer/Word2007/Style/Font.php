<?php
/**
 * This file is part of PHPWord - A pure PHP library for reading and writing
 * word processing documents.
 *
 * PHPWord is free software distributed under the terms of the GNU Lesser
 * General Public License version 3 as published by the Free Software Foundation.
 *
 * For the full copyright and license information, please read the LICENSE
 * file that was distributed with this source code. For the full list of
 * contributors, visit https://github.com/PHPOffice/PHPWord/contributors.
 *
 * @see         https://github.com/PHPOffice/PHPWord
 *
 * @copyright   2010-2018 PHPWord contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

namespace PhpOffice\PhpWord\Writer\Word2007\Style;

/**
 * Font style writer.
 *
 * @since 0.10.0
 */
class Font extends AbstractStyle
{
    /**
     * Is inline in element.
     *
     * @var bool
     */
    private $isInline = false;

    /**
     * Write style.
     */
    public function write()
    {
        $xmlWriter = $this->getXmlWriter();

        $isStyleName = $this->isInline && !is_null($this->style) && is_string($this->style);
        if ($isStyleName) {
            $xmlWriter->startElement('w:rPr');
            $xmlWriter->startElement('w:rStyle');
            $xmlWriter->writeAttribute('w:val', $this->style);
            $xmlWriter->endElement();
            $style = \PhpOffice\PhpWord\Style::getStyle($this->style);
            if ($style instanceof \PhpOffice\PhpWord\Style\Font) {
                $xmlWriter->writeElementIf($style->isRTL(), 'w:rtl');
                $xmlWriter->startElement('w:bidi');
                $xmlWriter->endElement();
            }
            $xmlWriter->endElement();
        } else {
            $this->writeStyle();
        }
    }

    /**
     * Set is inline.
     *
     * @param bool $value
     */
    public function setIsInline($value)
    {
        $this->isInline = $value;
    }

    /**
     * Write full style.
     */
    private function writeStyle()
    {
        $style = $this->getStyle();
        if (!$style instanceof \PhpOffice\PhpWord\Style\Font) {
            return;
        }

        $xmlWriter = $this->getXmlWriter();

        $xmlWriter->startElement('w:rPr');

        // Style name
        if (true === $this->isInline) {
            $styleName = $style->getStyleName();
            $xmlWriter->writeElementIf(null !== $styleName, 'w:rStyle', 'w:val', $styleName);
        }

        // Font name/family
        $font = $style->getName();
        $hint = $style->getHint();
        if (null !== $font) {
            $xmlWriter->startElement('w:rFonts');
            $xmlWriter->writeAttribute('w:ascii', $font);
            $xmlWriter->writeAttribute('w:hAnsi', $font);
            $xmlWriter->writeAttribute('w:eastAsia', $font);
            $xmlWriter->writeAttribute('w:cs', $font);
            $xmlWriter->writeAttributeIf(null !== $hint, 'w:hint', $hint);
            $xmlWriter->endElement();
        }

        //Language
        $language = $style->getLang();
        if (null != $language && (null !== $language->getLatin() || null !== $language->getEastAsia() || null !== $language->getBidirectional())) {
            $xmlWriter->startElement('w:lang');
            $xmlWriter->writeAttributeIf(null !== $language->getLatin(), 'w:val', $language->getLatin());
            $xmlWriter->writeAttributeIf(null !== $language->getEastAsia(), 'w:eastAsia', $language->getEastAsia());
            $xmlWriter->writeAttributeIf(null !== $language->getBidirectional(), 'w:bidi', $language->getBidirectional());
            //if bidi is not set but we are writing RTL, write the latin language in the bidi tag
            if ($style->isRTL() && null === $language->getBidirectional() && null !== $language->getLatin()) {
                $xmlWriter->writeAttribute('w:bidi', $language->getLatin());
            }
            $xmlWriter->endElement();
        }

        // Color
        $color = $style->getColor();
        $xmlWriter->writeElementIf(null !== $color, 'w:color', 'w:val', $color);

        // Size
        $size = $style->getSize();
        $xmlWriter->writeElementIf(null !== $size, 'w:sz', 'w:val', $size * 2);
        $xmlWriter->writeElementIf(null !== $size, 'w:szCs', 'w:val', $size * 2);

        // Bold, italic
        $xmlWriter->writeElementIf(null !== $style->isBold(), 'w:b', 'w:val', $this->writeOnOf($style->isBold()));
        $xmlWriter->writeElementIf(null !== $style->isBold(), 'w:bCs', 'w:val', $this->writeOnOf($style->isBold()));
        $xmlWriter->writeElementIf(null !== $style->isItalic(), 'w:i', 'w:val', $this->writeOnOf($style->isItalic()));
        $xmlWriter->writeElementIf(null !== $style->isItalic(), 'w:iCs', 'w:val', $this->writeOnOf($style->isItalic()));

        // Strikethrough, double strikethrough
        $xmlWriter->writeElementIf(null !== $style->isStrikethrough(), 'w:strike', 'w:val', $this->writeOnOf($style->isStrikethrough()));
        $xmlWriter->writeElementIf(null !== $style->isDoubleStrikethrough(), 'w:dstrike', 'w:val', $this->writeOnOf($style->isDoubleStrikethrough()));

        // Small caps, all caps
        $xmlWriter->writeElementIf(null !== $style->isSmallCaps(), 'w:smallCaps', 'w:val', $this->writeOnOf($style->isSmallCaps()));
        $xmlWriter->writeElementIf(null !== $style->isAllCaps(), 'w:caps', 'w:val', $this->writeOnOf($style->isAllCaps()));

        //Hidden text
        $xmlWriter->writeElementIf($style->isHidden(), 'w:vanish', 'w:val', $this->writeOnOf($style->isHidden()));

        // Underline
        $xmlWriter->writeElementIf('none' != $style->getUnderline(), 'w:u', 'w:val', $style->getUnderline());

        // Foreground-Color
        $xmlWriter->writeElementIf(null !== $style->getFgColor(), 'w:highlight', 'w:val', $style->getFgColor());

        // Superscript/subscript
        $xmlWriter->writeElementIf($style->isSuperScript(), 'w:vertAlign', 'w:val', 'superscript');
        $xmlWriter->writeElementIf($style->isSubScript(), 'w:vertAlign', 'w:val', 'subscript');

        // Spacing
        $xmlWriter->writeElementIf(null !== $style->getScale(), 'w:w', 'w:val', $style->getScale());
        $xmlWriter->writeElementIf(null !== $style->getSpacing(), 'w:spacing', 'w:val', $style->getSpacing());
        $xmlWriter->writeElementIf(null !== $style->getKerning(), 'w:kern', 'w:val', $style->getKerning() * 2);

        // noProof
        $xmlWriter->writeElementIf(null !== $style->isNoProof(), 'w:noProof', 'w:val', $this->writeOnOf($style->isNoProof()));

        // Background-Color
        $shading = $style->getShading();
        if (!is_null($shading)) {
            $styleWriter = new Shading($xmlWriter, $shading);
            $styleWriter->write();
        }

        // RTL
        if (true === $this->isInline) {
            $styleName = $style->getStyleName();
            $xmlWriter->writeElementIf(null === $styleName && $style->isRTL(), 'w:rtl');
        }

        // Position
        $xmlWriter->writeElementIf(null !== $style->getPosition(), 'w:position', 'w:val', $style->getPosition());

        $xmlWriter->endElement();
    }
}
