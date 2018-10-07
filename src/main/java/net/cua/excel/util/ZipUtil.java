package net.cua.excel.util;

/**
 * Created by guanquan.wang on 2017/10/13.
 */
import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardOpenOption;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;
import java.util.zip.*;

public class ZipUtil {
    public static final String suffix = ".zip";
    private ZipUtil() {}
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
                paths.addAll(Arrays.stream(src.toFile().listFiles()).map(File::toPath).collect(Collectors.toList()));
                while (i < paths.size()) {
                    if (Files.isDirectory(paths.get(i))) {
                        paths.addAll(Arrays.stream(paths.get(i).toFile().listFiles()).map(File::toPath).collect(Collectors.toList()));
                        // @FIX JDK BUG. => Files.list stream do not close resource
//                        paths.addAll(Files.list(paths.get(i)).collect(Collectors.toList()));
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
                zos.putNextEntry(new ZipEntry(name));
                Files.copy(paths.get(j), zos);
                zos.closeEntry();
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
            FileUtil.mkdir(destPath);
        }
        ZipInputStream zis = new ZipInputStream(stream);
        ZipEntry entry = zis.getNextEntry();
        while (entry != null) {
            Path sub = destPath.resolve(entry.getName());
            // Create parent
            if (!Files.exists(sub.getParent())) {
                FileUtil.mkdir(sub.getParent());
            }
            if (entry.isDirectory()) {
                FileUtil.mkdir(sub);
            } else {
                FileUtil.cp(zis, sub);
            }
            zis.closeEntry();
            entry = zis.getNextEntry();
        }

        zis.close();
        return destPath;
    }

}