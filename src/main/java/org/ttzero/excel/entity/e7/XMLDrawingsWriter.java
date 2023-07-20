/*
 * Copyright (c) 2017-2023, guanquan.wang@yandex.com All Rights Reserved.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */


package org.ttzero.excel.entity.e7;

import org.ttzero.excel.entity.IDrawingsWriter;
import org.ttzero.excel.entity.Picture;
import org.ttzero.excel.entity.Relationship;
import org.ttzero.excel.manager.Const;
import org.ttzero.excel.manager.RelManager;
import org.ttzero.excel.manager.TopNS;
import org.ttzero.excel.util.ExtBufferedWriter;
import org.ttzero.excel.util.FileUtil;
import org.ttzero.excel.util.StringUtil;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Arrays;

import static org.ttzero.excel.util.FileUtil.exists;

/**
 * Drawings writer(For Picture only)
 *
 * @author wangguanquan3 at 2023-03-07 09:09
 */
@TopNS(prefix = {"xdr", "a", "r"}, value = "wsDr"
    , uri = {"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
    , "http://schemas.openxmlformats.org/drawingml/2006/main"
    , Const.Relationship.RELATIONSHIP})
public class XMLDrawingsWriter implements IDrawingsWriter {
    protected Path path;
    protected ExtBufferedWriter bw;
    protected int size;
    protected RelManager relManager;
    /**
     * Store async pictures
     */
    protected Picture[] pictures;
    /**
     * Mark complete status
     * 0: free or complete
     * 1: wait
     */
    protected long[] bits;
    protected int countDown;

    private XMLDrawingsWriter() { }

    public XMLDrawingsWriter(Path path) {
        this.path = path;
        this.relManager = new RelManager();
        try {
            if (!exists(path.getParent())) {
                FileUtil.mkdir(path.getParent());
            }
            bw = new ExtBufferedWriter(Files.newBufferedWriter(path, StandardCharsets.UTF_8));
            bw.write("<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
        } catch (IOException e) {
            throw new RuntimeException("Create XMLDrawingsWriter error", e);
        }
    }

    @Override
    public void close() throws IOException {
        if (countDown > 0) {
            int counter = 30; // Loop 30 times (1 minutes)
            do {
                // Check status
                if (checkComplete() == 0) break;

                try {
                    Thread.sleep(2000);
                } catch (InterruptedException e) {
                    break;
                }
            } while (--counter >= 0);
        }
        // End tag
        bw.write("</xdr:wsDr>");
        relManager.write(path.getParent(), path.getFileName().toString());
        FileUtil.close(bw);
    }

    @Override
    public void writeTo(Path root) throws IOException { }

    static final String[] ANCHOR_PROPERTY = {"twoCell", "oneCell", "absolute"};

    @Override
    public void drawing(Picture picture) throws IOException {
        if (StringUtil.isEmpty(picture.picName)) return;
        Relationship picRel = relManager.add(new Relationship("../media/" + picture.picName, Const.Relationship.IMAGE));
        size++;

        bw.write("<xdr:twoCellAnchor editAs=\"");
        bw.write(ANCHOR_PROPERTY[picture.property & 3]);
        // Default editAs="twoCell"
        bw.write("\">");

        // From
        bw.write("<xdr:from><xdr:col>");
        bw.writeInt(picture.col);
        bw.write("</xdr:col><xdr:colOff>");
        bw.writeInt((picture.padding >>> 24) * 12700);
        bw.write("</xdr:colOff><xdr:row>");
        bw.writeInt(picture.row - 1);
        bw.write("</xdr:row><xdr:rowOff>");
        bw.writeInt(((picture.padding >>> 16) & 0xFF) * 12700);
        bw.write("</xdr:rowOff></xdr:from>");

        // TO
        bw.write("<xdr:to><xdr:col>");
        bw.writeInt(picture.col + 1);
        bw.write("</xdr:col><xdr:colOff>");
        bw.writeInt(-((picture.padding >>> 8) & 0xFF) * 12700);
        bw.write("</xdr:colOff><xdr:row>");
        bw.writeInt(picture.row);
        bw.write("</xdr:row><xdr:rowOff>");
        bw.writeInt(-(picture.padding & 0xFF) * 12700);
        bw.write("</xdr:rowOff></xdr:to>");

        // Picture
        bw.write("<xdr:pic><xdr:nvPicPr><xdr:cNvPr id=\"");
        bw.writeInt(size);
        bw.write("\" name=\"Picture ");
        bw.writeInt(size);
        bw.write("\"/><xdr:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></xdr:cNvPicPr></xdr:nvPicPr>");
        bw.write("<xdr:blipFill><a:blip r:embed=\"");
        bw.write(picRel.getId());
        bw.write("\"/>");

        bw.write("<a:stretch><a:fillRect/></a:stretch></xdr:blipFill>");
        bw.write("<xdr:spPr><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></xdr:spPr>");
//        bw.write("<xdr:spPr><a:prstGeom prst=\"roundRect\"><a:avLst><a:gd name=\"adj\" fmla=\"val 20594\"/></a:avLst></a:prstGeom></xdr:spPr>");
        // End Picture
        bw.write("</xdr:pic><xdr:clientData/></xdr:twoCellAnchor>");
    }

    @Override
    public void asyncDrawing(Picture picture) throws IOException {
        if (pictures == null) {
            bits = new long[2];
            pictures = new Picture[bits.length << 6];
        }
        int freeIndex = getFreeIndex(bits);
        // Grow and copy
        if (freeIndex < 0) {
            freeIndex = bits.length << 6;
            bits = Arrays.copyOf(bits, bits.length + 2);
            pictures = Arrays.copyOf(pictures, bits.length << 6);
        }
        // Write file if current location is completed
        else if (pictures[freeIndex] != null) {
            drawing(pictures[freeIndex]);
            countDown--;
        }

        picture.idx = freeIndex;
        pictures[freeIndex] = picture;
        markIndex(bits, freeIndex);
        countDown++;
    }

    @Override
    public void complete(Picture picture) {
        if (bits == null) return;
        freeIndex(bits, picture.idx);
    }

    protected int checkComplete() throws IOException {
        while (countDown > 0) {
            int i = getFreeIndex(bits);
            // None complete picture
            if (i < 0) break;
            Picture p = pictures[i];
            // Overflow
            if (p == null) break;
            // Write picture
            drawing(p);
            // The completed position is marked as 1 to prevent further acquisition
            markIndex(bits, i);
            countDown--;
        }
        return countDown;
    }

    public static int getFreeIndex(long[] bits) {
        int i = 0, idx = 64;
        for (; i < bits.length && (idx = Long.numberOfTrailingZeros(Long.highestOneBit(~bits[i]))) == 64; i++);
        return idx < 64 ? (i << 6) + (64 - idx - 1) : -1;
    }

    /**
     * Mark the bits at the specified position as 1
     */
    public static void markIndex(long[] bits, int idx) {
        bits[idx >> 6] |= 1L << (63 - (idx - (idx >>> 6 << 6)));
    }

    /**
     * Mark the bits at the specified position as 0
     */
    public static void freeIndex(long[] bits, int idx) {
        bits[idx >> 6] &= ~(1L << (63 - (idx - (idx >>> 6 << 6))));
    }
}
