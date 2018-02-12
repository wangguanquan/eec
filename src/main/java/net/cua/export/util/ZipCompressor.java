package net.cua.export.util;

/**
 * Created by guanquan.wang on 2017/10/13.
 */
import java.io.*;
import java.util.zip.CRC32;
import java.util.zip.CheckedOutputStream;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;


public class ZipCompressor {
    static final int BUFFER = 8192;

    private File zipFile;

    public ZipCompressor(String pathName) {
        if (pathName.lastIndexOf(".zip") != pathName.length() - 4) {
            pathName += ".zip";
        }
        zipFile = new File(pathName);
    }

    /**
     * 从当前文件夹向下压缩包含当前目录
     * @param srcPathName
     */
    public File compress(String srcPathName) {
        File file = new File(srcPathName);
        if (!file.exists())
            throw new RuntimeException(srcPathName + " not exists.");
        ZipOutputStream out = null;
        try {
            out = new ZipOutputStream(new CheckedOutputStream(new FileOutputStream(zipFile), new CRC32()));
            String basedir = "";
            compress(file, out, basedir);
        } catch (IOException e) {
            throw new RuntimeException(e);
        } finally {
            if (out != null) {
                try {
                    out.close();
                } catch (IOException e) {
                }
            }
        }
        return zipFile;
    }

    /**
     * 从当前文件夹向下压缩不含当前目录
     * @param file
     */
    public File compressSubs(File file) {
        if (file.isDirectory()) {
            ZipOutputStream out = null;
            try {
                out = new ZipOutputStream(new CheckedOutputStream(new FileOutputStream(zipFile), new CRC32()));
                String basedir = "";

                File[] files = file.listFiles();
                if (files != null) {
                    for (File f : files) {
                        compress(f, out, basedir);
                    }
                }
            } catch (IOException e) {
                throw new RuntimeException(e);
            } finally {
                if (out != null) {
                    try {
                        out.close();
                    } catch (IOException e) {
                    }
                }
            }
        } else {
            compress(file.getPath());
        }
        return zipFile;
    }

    private void compress(File file, ZipOutputStream out, String basedir) {
        /* 判断是目录还是文件 */
        if (file.isDirectory()) {
            compressDirectory(file, out, basedir);
        } else {
            compressFile(file, out, basedir);
        }
    }

    /** 压缩一个目录 */
    private void compressDirectory(File dir, ZipOutputStream out, String basedir) {
        File[] files = dir.listFiles();
        if (files == null) return;
        for (int i = 0; i < files.length; i++) {
            /* 递归 */
            compress(files[i], out, basedir + dir.getName() + File.separatorChar);
        }
    }

    /** 压缩一个文件 */
    private void compressFile(File file, ZipOutputStream out, String basedir) {
        BufferedInputStream bis = null;
        try {
            bis = new BufferedInputStream(new FileInputStream(file));
            ZipEntry entry = new ZipEntry(basedir + file.getName());
            out.putNextEntry(entry);
            int count;
            byte data[] = new byte[BUFFER];
            while ((count = bis.read(data, 0, BUFFER)) != -1) {
                out.write(data, 0, count);
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        } finally {
            if (bis != null) {
                try {
                    bis.close();
                } catch (IOException e) {
                }
            }
        }
    }
}