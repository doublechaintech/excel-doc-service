package com.doublechaintech.xls.service;

public class Response {
  private int code;
  private String message;
  private String output;
  private long took;

  public String getOutput() {
    return output;
  }

  public void setOutput(String pOutput) {
    output = pOutput;
  }

  public long getTook() {
    return took;
  }

  public void setTook(long pTook) {
    took = pTook;
  }

  public int getCode() {
    return code;
  }

  public void setCode(int pCode) {
    code = pCode;
  }

  public String getMessage() {
    return message;
  }

  public void setMessage(String pMessage) {
    message = pMessage;
  }
}
