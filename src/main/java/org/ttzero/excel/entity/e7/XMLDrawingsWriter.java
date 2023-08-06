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

import org.ttzero.excel.drawing.Angle;
import org.ttzero.excel.drawing.Bevel;
import org.ttzero.excel.drawing.Camera;
import org.ttzero.excel.drawing.Fill;
import org.ttzero.excel.drawing.Glow;
import org.ttzero.excel.drawing.Guide;
import org.ttzero.excel.drawing.LightRig;
import org.ttzero.excel.drawing.Outline;
import org.ttzero.excel.drawing.PictureEffect;
import org.ttzero.excel.drawing.Reflection;
import org.ttzero.excel.drawing.Scene3D;
import org.ttzero.excel.drawing.Shadow;
import org.ttzero.excel.drawing.Shape3D;
import org.ttzero.excel.drawing.ShapeType;
import org.ttzero.excel.entity.IDrawingsWriter;
import org.ttzero.excel.entity.Picture;
import org.ttzero.excel.entity.Relationship;
import org.ttzero.excel.entity.style.ColorIndex;
import org.ttzero.excel.manager.Const;
import org.ttzero.excel.manager.RelManager;
import org.ttzero.excel.manager.TopNS;
import org.ttzero.excel.util.ExtBufferedWriter;
import org.ttzero.excel.util.FileUtil;
import org.ttzero.excel.util.StringUtil;

import java.awt.Color;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Arrays;

import static org.ttzero.excel.util.FileUtil.exists;

/**
 * Drawings writer(For Picture only)
 *
 * @author guanquan.wng at 2023-03-07 09:09
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
        if (bw == null) return;
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
        bw = null;
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
        bw.writeInt((picture.padding & 0xFF) * 12700);
        bw.write("</xdr:colOff><xdr:row>");
        bw.writeInt(picture.row - 1);
        bw.write("</xdr:row><xdr:rowOff>");
        bw.writeInt((picture.padding >>> 24) * 12700);
        bw.write("</xdr:rowOff></xdr:from>");

        // TO
        bw.write("<xdr:to><xdr:col>");
        bw.writeInt(picture.col + 1);
        bw.write("</xdr:col><xdr:colOff>");
        bw.writeInt(-((picture.padding >>> 16) & 0xFF) * 12700);
        bw.write("</xdr:colOff><xdr:row>");
        bw.writeInt(picture.row);
        bw.write("</xdr:row><xdr:rowOff>");
        bw.writeInt(-((picture.padding >>> 8) & 0xFF) * 12700);
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
        bw.write("<xdr:spPr>");

        // Revolve
        if (picture.revolve != 0) {
            bw.write("<a:xfrm rot=\""); bw.writeInt(60000 * picture.revolve); bw.write("\"/>");
        }

        // Picture Effects
        if (picture.effect == null) {
            // Default geometry: rect
            bw.write("<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>");
        }
        // Write effects
        else {
            writeEffects(picture);
        }

        bw.write("</xdr:spPr>");
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
     *
     * @param bits bit array
     * @param idx index
     */
    public static synchronized void markIndex(long[] bits, int idx) {
        bits[idx >> 6] |= 1L << (63 - (idx - (idx >>> 6 << 6)));
    }

    /**
     * Mark the bits at the specified position as 0
     *
     * @param bits bit array
     * @param idx index
     */
    public static synchronized void freeIndex(long[] bits, int idx) {
        bits[idx >> 6] &= ~(1L << (63 - (idx - (idx >>> 6 << 6))));
    }

    protected void writeEffects(Picture pict) throws IOException {
        PictureEffect effect = pict.effect;
        // Geometry
        bw.write("<a:prstGeom prst=\"");
        bw.write(effect.geometry != null ? effect.geometry.name() : ShapeType.rect.name());
        if (effect.geometryAdjustValueList != null && !effect.geometryAdjustValueList.isEmpty()) {
            bw.write("\"><a:avLst>");
            for (Guide guide : effect.geometryAdjustValueList) {
                bw.write("<a:gd name=\"");
                bw.write(guide.name);
                bw.write("\" fmla=\"");
                bw.write(guide.fmla);
                bw.write("\"/>");
            }
            bw.write("</a:avLst></a:prstGeom>");
        } else bw.write("\"><a:avLst/></a:prstGeom>");

        // Fill
        if (effect.fill != null) attachFill(effect.fill);

        // Outline
        if (effect.outline != null && effect.outline.width > 0.) attachOutline(effect.outline);

        boolean hasShadow = effect.shadow != null && effect.shadow.size > 0
            , hasInnerShadow = effect.innerShadow != null
            , hasReflection = effect.reflection != null
            , hasSoftEdges = effect.softEdges > 0
            , hasGlow = effect.glow != null;

        // Picture Effects
        if (hasShadow || hasInnerShadow || hasReflection || hasSoftEdges || hasGlow) {
            bw.write("<a:effectLst>");

            // Glow
            if (hasGlow) attachGlow(effect.glow);

            // Shadow
            if (hasShadow) attachShadow(effect.shadow, "outerShdw");
            if (hasInnerShadow) attachShadow(effect.innerShadow, "innerShdw");

            // Reflection
            if (hasReflection) attachReflection(effect.reflection);

            // Soft Edges
            if (hasSoftEdges) attachSoftEdges(effect.softEdges);

            bw.write("</a:effectLst>");
        }

        // 3D Scene
        if (effect.scene3D != null && effect.scene3D.camera != null) attachScene3d(effect.scene3D);

        // 3D Shape
        if (effect.shape3D != null) attachShape3D(effect.shape3D);
    }

    protected void attachShadow(Shadow shadow, String tag) throws IOException {
        bw.write("<a:"); bw.write(tag);
        int n = (int) (shadow.blur % 101 * 12700);
        if (n > 0) {
            bw.write(" blurRad=\""); bw.writeInt(n);
        }
        n = (int) (shadow.dist % 201 * 12700) / 100 * 100;
        if (n > 0) {
            bw.write("\" dist=\""); bw.writeInt(n);
        }
        n = shadow.direction % 361 * 60000;
        if (n > 0) {
            bw.write("\" dir=\""); bw.writeInt(n);
        }
        boolean hs = shadow.sx != 0. || shadow.sy != 0.;
        if (hs) {
            if (shadow.sx != 0.) {
                bw.write("\" sx=\"");
                bw.writeInt(((int) (shadow.sx % 201 * 1000)));
            }
            if (shadow.sy != 0.) {
                bw.write("\" sy=\"");
                bw.writeInt(((int) (shadow.sy % 201 * 1000)));
            }
        } else if (shadow.size != 100.) {
            bw.write("\" sx=\"");
            bw.writeInt(n = ((int) (shadow.size % 201 * 1000)));
            bw.write("\" sy=\"");
            bw.writeInt(n);
        }
        if (shadow.kx > 0) {
            bw.write("\" kx=\"");
            bw.writeInt(((int) (shadow.kx % 361 * 60000)));
        }
        if (shadow.ky > 0) {
            bw.write("\" ky=\"");
            bw.writeInt(((int) (shadow.ky % 361 * 60000)));
        }
        if (shadow.angle != null) {
            bw.write("\" algn=\"");
            bw.write(shadow.angle.shotName);
        }
        if ("outerShdw".equals(tag)) {
            bw.write("\" rotWithShape=\"");
            bw.writeInt(shadow.rotWithShape);
        }
        bw.write("\"><a:srgbClr val=\""); bw.write(ColorIndex.toRGB(shadow.color));
        n = (100 - shadow.alpha % 101) * 1000;
        if (n != 100000) {
            bw.write("\"><a:alpha val=\"");
            bw.writeInt(n);
            bw.write("\"/></a:srgbClr>");
        } else bw.write("\"/>");
        bw.write("</a:"); bw.write(tag); bw.write(">");
    }

    protected void attachReflection(Reflection reflection) throws IOException {
        bw.write("<a:reflection blurRad=\"");
        bw.writeInt((int) (reflection.blur % 101 * 12700));
        bw.write("\" stA=\"");
        bw.writeInt((100 - reflection.alpha % 101) * 1000);
        if (reflection.dist > 0) {
            bw.write("\" dist=\"");
            bw.writeInt((int) (reflection.dist % 101 * 12700));
        }
        bw.write("\" endPos=\"");
        bw.writeInt(reflection.size % 101 * 1000);
        bw.write("\" dir=\"");
        bw.writeInt(reflection.direction % 361 * 60000);
        bw.write("\" sy=\"-100000\" algn=\"bl\" rotWithShape=\"0\"/>");
    }

    protected void attachGlow(Glow glow) throws IOException {
        int red = (int) (glow.dist % 151 * 12700) / 100 * 100;
        bw.write("<a:glow");
        if (red > 0) {
            bw.write(" rad=\"");
            bw.writeInt(red);
            bw.write('\"');
        }
        bw.write("><a:srgbClr val=\"");
        Color color = glow.color != null ? glow.color : Color.WHITE;
        bw.write(ColorIndex.toRGB(color));
        int t = 100 - glow.alpha % 101;
        if (t != 100) {
            bw.write("\"><a:alpha val=\"");
            bw.writeInt(t * 1000);
            bw.write("\"/></a:srgbClr></a:glow>");
        } else bw.write("\"/></a:glow>");
    }

    protected void attachSoftEdges(double softEdges) throws IOException {
        bw.write("<a:softEdge rad=\"");
        bw.writeInt((int) (softEdges % 101 * 12700) / 100 * 100);
        bw.write("\"/>");
    }

    protected void attachFill(Fill fill) throws IOException {
        if (fill instanceof Fill.SolidFill) {
            Fill.SolidFill solidFill = (Fill.SolidFill) fill;
            if (solidFill.color == null) return;
            bw.write("<a:solidFill><a:srgbClr val=\"");
            bw.write(ColorIndex.toRGB(solidFill.color));
            bw.write("\">");
            if (solidFill.shade > 0) {
                bw.write("<a:shade val=\"");
                bw.writeInt(solidFill.shade * 1000);
                bw.write("\"/>");
            }
            int t = 100 - solidFill.alpha % 101;
            if (t != 100) {
                bw.write("<a:alpha val=\"");
                bw.writeInt(t * 1000);
                bw.write("\"/>");
            }
            bw.write("</a:srgbClr></a:solidFill>");
        }
        // Pattern fill
        else if (fill instanceof Fill.PatternFill) {
            // TODO
        }
        // Gradient fill
        else if (fill instanceof Fill.GradientFill) {
            // TODO
        }
    }

    protected void attachOutline(Outline ln) throws IOException {
        bw.write("<a:ln w=\"");
        bw.writeInt((int) (ln.width * 12700));
        bw.write("\" cap=\"");
        bw.write(ln.cap != null ? ln.cap.shotName : Outline.Cap.SQUARE.shotName);
        if (ln.cmpd != null) {
            bw.write("\" cmpd=\"");
            bw.write(ln.cmpd.shotName);
        }
        bw.write("\">");
        // Support 'solidFill' only
        bw.write("<a:solidFill><a:srgbClr val=\"");
        bw.write(ln.color != null ? ColorIndex.toRGB(ln.color) : "FFFFFF");
        bw.write("\"/></a:solidFill>");
        if (ln.dash != null) {
            bw.write("<a:prstDash val=\""); bw.write(ln.dash.name()); bw.write("\"/>");
        }
        if (ln.joinType != null) {
            bw.write("<a:"); bw.write(ln.joinType.name());
            if (ln.joinType == Outline.JoinType.miter) {
                bw.write(" lim=\"");
                bw.writeInt(ln.miterLimit > 0 ? ((int) (ln.miterLimit * 1000)) : 800000);
                bw.write("\"/>");
            } else bw.write("/>");
        }
        bw.write("</a:ln>");
    }

    protected void attachScene3d(Scene3D scene) throws IOException {
        bw.write("<a:scene3d>");
        Camera camera = scene.camera;
        bw.write("<a:camera prst=\"");
        bw.write(camera.presetCamera != null ? camera.presetCamera.name() : Camera.PresetCamera.orthographicFront.name());
        if (camera.fov > 0.) {
            bw.write("\" fov=\"");
            bw.writeInt((int) (camera.fov % 181 * 60000));
        }
        if (camera.zoom > 0) {
            bw.write("\" zoom=\"");
            bw.writeInt(camera.zoom);
            bw.write("%");
        }
        if (camera.latitude > 0. || camera.longitude > 0. || camera.revolution > 0.) {
            bw.write("\"><a:rot lat=\"");
            bw.writeInt((int) (camera.latitude * 60000));
            bw.write("\" lon=\"");
            bw.writeInt((int) (camera.longitude * 60000));
            bw.write("\" rev=\"");
            bw.writeInt((int) (camera.revolution * 60000));
            bw.write("\"/></a:camera>");
        } else bw.write("\"/>");
        LightRig lightRig = scene.lightRig;
        if (lightRig != null && lightRig.rig != null) {
            bw.write("<a:lightRig rig=\"");
            bw.write(lightRig.rig.name());
            bw.write("\" dir=\"");
            bw.write(lightRig.angle != null ? lightRig.angle.shotName : Angle.TOP.shotName);
            bw.write("\">");
            if (lightRig.latitude > 0. || lightRig.longitude > 0. || lightRig.revolution > 0.) {
                bw.write("<a:rot lat=\"");
                bw.writeInt((int) (lightRig.latitude * 60000));
                bw.write("\" lon=\"");
                bw.writeInt((int) (lightRig.longitude * 60000));
                bw.write("\" rev=\"");
                bw.writeInt((int) (lightRig.revolution * 60000));
                bw.write("\"/>");
            }
            bw.write("</a:lightRig>");
        }
        bw.write("</a:scene3d>");
    }

    protected void attachShape3D(Shape3D shape) throws IOException {
        bw.write("<a:sp3d");
        if (shape.contourWidth > 0.) writeShapeProp(shape.contourWidth, "contourW");
        if (shape.extrusionHeight > 0.) writeShapeProp(shape.extrusionHeight, "extrusionH");
        if (shape.material != null) {
            bw.write(" prstMaterial=\"");
            bw.write(shape.material.name());
            bw.write("\"");
        }
        bw.write(">");
        if (shape.bevelBottom != null) writeBevel(shape.bevelBottom, 'B');
        if (shape.bevelTop != null) writeBevel(shape.bevelTop, 'T');
        if (shape.contourColor != null) {
            bw.write("<a:contourClr><a:srgbClr val=\"");
            bw.write(ColorIndex.toRGB(shape.contourColor));
            bw.write("\"/></a:contourClr>");
        }
        if (shape.extrusionColor != null) {
            bw.write("<a:extrusionClr><a:srgbClr val=\"");
            bw.write(ColorIndex.toRGB(shape.extrusionColor));
            bw.write("\"/></a:extrusionClr>");
        }
        bw.write("</a:sp3d>");
    }

    protected void writeBevel(Bevel bevel, char angle) throws IOException {
        bw.write("<a:bevel"); bw.write(angle); bw.write(" w=\"");
        bw.writeInt((int) (bevel.width * 12700));
        bw.write("\" h=\"");
        bw.writeInt((int) (bevel.height * 12700));
        if (bevel.prst != null) {
            bw.write("\" prst=\"");
            bw.write(bevel.prst.name());
        }
        bw.write("\"/>");
    }

    private void writeShapeProp(double v, String tag) throws IOException {
        bw.write(" "); bw.write(tag); bw.write("=\"");
        bw.writeInt((int) (v * 12700));
        bw.write("\"");
    }
}
