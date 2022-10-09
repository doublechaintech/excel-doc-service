package com.doublechaintech.xls;

import java.io.IOException;
import java.io.OutputStream;
import java.util.List;
import java.util.stream.Stream;

public interface BlockWriter {
  default void append(Stream<Block> blockStream) {
    if (blockStream == null) {
      return;
    }
    blockStream.forEach(this::append);
  }

  default void append(List<Block> blocks) {
    if (blocks == null) {
      return;
    }
    blocks.forEach(this::append);
  }

  void append(Block pBl);

  void write(OutputStream out) throws IOException;
}
