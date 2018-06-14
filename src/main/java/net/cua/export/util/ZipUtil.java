package net.cua.export.util;

/**
 * Created by guanquan.wang on 2017/10/13.
 */
import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardOpenOption;
import java.nio.file.attribute.PosixFilePermissions;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.*;

public class ZipUtil {
    static final int BUFFER = 8192;
    public static final String suffix = ".zip";

    /**
     * zip files exclude root path
     * command: zip destPath srcPath1 srcPath2 ...
     * @param destPath
     * @param srcPath
     * @return
     * @throws IOException
     */
    public static Path zip(Path destPath, Path ... srcPath) throws IOException {
        return zip(destPath, true, srcPath);
    }

    /**
     * zip files include root path
     * command: zip destPath srcPath1 srcPath2 ...
     * @param destPath
     * @param compressRoot
     * @param srcPath
     * @return
     * @throws IOException
     */
    public static Path zip(Path destPath, boolean compressRoot, Path ... srcPath) throws IOException {
        if (!destPath.toString().endsWith(suffix)) {
            destPath = Paths.get(destPath.toString() + suffix);
        }
        ZipOutputStream zos = new ZipOutputStream(new CheckedOutputStream(
                Files.newOutputStream(destPath, StandardOpenOption.CREATE), new CRC32()));
        List<Path> paths = new ArrayList<>();
        int i = 0, index = 0;
        int[] array = new int[srcPath.length];
        for (Path src : srcPath) {
            if (Files.isDirectory(src)) {
                Files.list(src).forEach(path -> paths.add(path));
                while (i < paths.size()) {
                    if (Files.isDirectory(paths.get(i))) {
                        Files.list(paths.get(i)).forEach(path -> paths.add(path));
                    }
                    i++;
                }
            } else {
                paths.add(src);
                i++;
            }
            array[index++] = i;
        }

        index = 0;
        byte buff[] = new byte[BUFFER];
        Path basePath = compressRoot ? srcPath[index].getParent() : srcPath[index];
        for (int j = 0, len = basePath.toString().length(); j < i; j++) {
            if (Files.isDirectory(paths.get(j))) continue;
            if (j < array[index]) {
                String name;
                if (paths.get(j).equals(srcPath[index])) {
                    name = paths.get(j).toString().substring(paths.get(j).getParent().toString().length() + 1);
                } else {
                    name = paths.get(j).toString().substring(len + 1);
                }
                ZipEntry entry = new ZipEntry(name);
                try (InputStream is = Files.newInputStream(paths.get(j), StandardOpenOption.READ)) {
                    zos.putNextEntry(entry);
                    int count;
                    while ((count = is.read(buff)) != -1) {
                        zos.write(buff, 0, count);
                    }
                    zos.closeEntry();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            } else {
                basePath = compressRoot ? srcPath[++index].getParent() : srcPath[++index];
                len = basePath.toString().length();
                j--;
            }
        }

        zos.close();
        return destPath;
    }

    /**
     * unzip file to descPath
     * @param stream
     * @param destPath
     * @return
     * @throws IOException
     */
    public static Path unzip(InputStream stream, Path destPath) throws IOException {
        if (!Files.isDirectory(destPath)) {
            Files.createDirectories(destPath);
        }
        ZipInputStream zis = new ZipInputStream(stream);
        ZipEntry entry = zis.getNextEntry();
        byte[] buff = new byte[BUFFER];
        while (entry != null) {
            Path sub = destPath.resolve(entry.getName());
            // Create parent
            if (!Files.exists(sub.getParent())) {
                Files.createDirectories(sub.getParent()
                        , PosixFilePermissions.asFileAttribute(PosixFilePermissions.fromString("rwxr-x---")));
            }
            if (entry.isDirectory()) {
                Files.createDirectory(sub
                        , PosixFilePermissions.asFileAttribute(PosixFilePermissions.fromString("rwxr-x---")));
            } else {
                OutputStream outputStream = Files.newOutputStream(sub, StandardOpenOption.CREATE);
                int len;
                while ((len = zis.read(buff)) > 0) {
                    outputStream.write(buff, 0, len);
                }
                FileUtil.close(outputStream);
            }
            zis.closeEntry();
            entry = zis.getNextEntry();
        }

        zis.close();
        return destPath;
    }
}