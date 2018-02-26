package net.cua.export.entity.e7;

import net.cua.export.entity.WaterMark;
import net.cua.export.util.ExtBufferedWriter;
import org.apache.log4j.Logger;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;

/**
 * Created by guanquan.wang at 2018-01-29 16:05
 */
public class EmptySheet extends Sheet {
    private Logger logger = Logger.getLogger(this.getClass().getName());

    public EmptySheet(Workbook workbook, String name, HeadColumn ... headColumns) {
        super(workbook, name, headColumns);
    }

    public EmptySheet(Workbook workbook, String name, WaterMark waterMark, HeadColumn ... headColumns) {
        super(workbook, name, waterMark, headColumns);
    }

    @Override
    public void close() {
        ;
    }

    @Override
    public void writeTo(Path xl) throws IOException {
        Path worksheets = xl.resolve("worksheets");
        if (!Files.exists(worksheets)) {
            Files.createDirectory(worksheets);
        }
        String name = getFileName();
//        logger.info(getName() + " | " + name);


        File sheetFile = worksheets.resolve(name).toFile();

        // write date
        try (ExtBufferedWriter bw = new ExtBufferedWriter(new OutputStreamWriter(new FileOutputStream(sheetFile), StandardCharsets.UTF_8))) {
            // Write header
            writeBefore(bw);
            // Main data
            // write ten empty rows
            for (int i = 0; i < 10; i++) {
                writeEmptyRow(bw);
            }

            // Write foot
            writeAfter(bw);

        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            close();
        }

        // resize columns
        boolean resize = false;
        for  (HeadColumn hc : headColumns) {
            if (hc.getWidth() > 0.000001) {
                resize = true;
                break;
            }
        }
        if (getAutoSize() == 1 || resize) {
            autoColumnSize(sheetFile);
        }

        // relationship
        relManager.write(worksheets, name);
    }
}
