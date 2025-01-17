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

namespace PhpOffice\PhpWord\Writer\Word2007\Element;

use PhpOffice\PhpWord\Element\TrackChange;

/**
 * Text element writer.
 *
 * @since 0.10.0
 */
class Text extends AbstractElement
{
    /**
     * Write text element.
     */
    public function write()
    {
        $xmlWriter = $this->getXmlWriter();
        $element = $this->getElement();
        if (!$element instanceof \PhpOffice\PhpWord\Element\Text) {
            return;
        }

        $this->startElementP();

        $xmlWriter->startElement('w:pPr');
        $xmlWriter->startElement('w:bidi');
        $xmlWriter->endElement();
        $xmlWriter->endElement();

        $this->writeOpeningTrackChange();

        $xmlWriter->startElement('w:r');

        $this->writeFontStyle();

        $xmlWriter->startElement('w:rPr'); //Added
        $xmlWriter->startElement('w:rtl'); //Added
        $xmlWriter->endElement(); //Added
        $xmlWriter->endElement(); //Added
        $textElement = 'w:t';
        //'w:delText' in case of deleted text
        $changed = $element->getTrackChange();
        if (null != $changed && TrackChange::DELETED == $changed->getChangeType()) {
            $textElement = 'w:delText';
        }
        $xmlWriter->startElement($textElement);

        $xmlWriter->writeAttribute('xml:space', 'preserve');
        $this->writeText($this->getText($element->getText()));
        $xmlWriter->endElement();
        $xmlWriter->endElement(); // w:r

        $this->writeClosingTrackChange();

        $this->endElementP(); // w:p
    }

    /**
     * Write opening of changed element.
     */
    protected function writeOpeningTrackChange()
    {
        $changed = $this->getElement()->getTrackChange();
        if (null == $changed) {
            return;
        }

        $xmlWriter = $this->getXmlWriter();

        if ((TrackChange::INSERTED == $changed->getChangeType())) {
            $xmlWriter->startElement('w:ins');
        } elseif (TrackChange::DELETED == $changed->getChangeType()) {
            $xmlWriter->startElement('w:del');
        }
        $xmlWriter->writeAttribute('w:author', $changed->getAuthor());
        if (null != $changed->getDate()) {
            $xmlWriter->writeAttribute('w:date', $changed->getDate()->format('Y-m-d\TH:i:s\Z'));
        }
        $xmlWriter->writeAttribute('w:id', $this->getElement()->getElementId());
    }

    /**
     * Write ending.
     */
    protected function writeClosingTrackChange()
    {
        $changed = $this->getElement()->getTrackChange();
        if (null == $changed) {
            return;
        }

        $xmlWriter = $this->getXmlWriter();

        $xmlWriter->endElement(); // w:ins|w:del
    }
}
