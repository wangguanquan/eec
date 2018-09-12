package net.cua.export.processor;

import java.nio.file.Path;

@FunctionalInterface
public interface DownProcessor {
    void build(Path path);
}
