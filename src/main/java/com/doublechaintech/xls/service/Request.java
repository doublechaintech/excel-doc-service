package com.doublechaintech.xls.service;

import com.doublechaintech.xls.Block;

import java.util.List;

public class Request {
  private String template;
  private List<Block> blocks;

  public String getTemplate() {
    return template;
  }

  public void setTemplate(String pTemplate) {
    template = pTemplate;
  }

  public List<Block> getBlocks() {
    return blocks;
  }

  public void setBlocks(List<Block> pBlocks) {
    blocks = pBlocks;
  }
}
