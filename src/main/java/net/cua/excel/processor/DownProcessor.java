package net.cua.excel.processor;

import java.nio.file.Path;

@FunctionalInterface
public interface DownProcessor {
    void build(Path path);
}
