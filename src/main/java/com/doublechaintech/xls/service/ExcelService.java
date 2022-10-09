package com.doublechaintech.xls.service;

import cn.hutool.core.codec.Base64;
import cn.hutool.core.io.FileUtil;
import cn.hutool.core.util.ObjectUtil;
import com.doublechaintech.xls.Block;
import com.doublechaintech.xls.XlsWriter;

import javax.ws.rs.GET;
import javax.ws.rs.POST;
import javax.ws.rs.Path;
import javax.ws.rs.Produces;
import javax.ws.rs.core.MediaType;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@Path("/xls")
public class ExcelService {

  @POST
  @Path("/generate")
  @Produces(MediaType.APPLICATION_JSON)
  public Response generate(Request request) {
    long startTime = System.currentTimeMillis();
    if (ObjectUtil.isEmpty(request)) {
      return error(1, "缺少参数request", startTime);
    }
    String template = request.getTemplate();
    List<Block> blocks = request.getBlocks();
    if (ObjectUtil.isEmpty(blocks)) {
      return error(2, "缺少blocks", startTime);
    }

    XlsWriter xlsWriter = new XlsWriter(template);
    xlsWriter.append(blocks);

    ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
    try {
      xlsWriter.write(outputStream);
    } catch (IOException pE) {
      pE.printStackTrace();
    }
    Response response = new Response();
    response.setOutput(Base64.encode(outputStream.toByteArray()));
    response.setTook(System.currentTimeMillis() - startTime);
    return response;
  }

  private Response error(int code, String message, long startTime) {
    Response response = new Response();
    response.setCode(code);
    response.setMessage(message);
    response.setTook(System.currentTimeMillis() - startTime);
    return response;
  }

  @GET
  @Path("/test")
  public javax.ws.rs.core.Response test() {
    Request request = new Request();
    List<Block> blocks = new ArrayList<>();
    byte[] templates =
        FileUtil.readBytes(
            new File(
                "/Users/jackytian/git/excel-doc-service/src/main/java/com/doublechaintech/xls/service/test.xlsx"));
    Block block = new Block();
    block.setBottom(1);
    block.setTop(1);
    block.setLeft(1);
    block.setRight(1);
    block.setValue("Hello");
    blocks.add(block);

    block = new Block();
    block.setBottom(1);
    block.setTop(1);
    block.setLeft(2);
    block.setRight(2);
    block.setValue("World");
    blocks.add(block);
    request.setBlocks(blocks);

    request.setTemplate(Base64.encode(templates));

    Response response = generate(request);
    if (response.getCode() == 0) {
      javax.ws.rs.core.Response.ResponseBuilder responseBuilder =
          javax.ws.rs.core.Response.ok(
              Base64.decode(response.getOutput()), MediaType.APPLICATION_OCTET_STREAM);
      responseBuilder.header("Content-Disposition", "attachment;filename=test.xlsx");
      return responseBuilder.build();
    }
    return javax.ws.rs.core.Response.serverError().build();
  }
}
