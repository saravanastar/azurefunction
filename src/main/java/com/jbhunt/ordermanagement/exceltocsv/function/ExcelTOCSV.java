package com.jbhunt.ordermanagement.exceltocsv.function;

import com.jbhunt.ordermanagement.exceltocsv.function.handler.ExcelToCSVHanlder;
import com.microsoft.azure.functions.annotation.*;
import com.microsoft.azure.functions.*;

import java.io.File;

/**
 * Azure Functions with Azure Blob trigger.
 */
public class ExcelTOCSV {

    private static String destination = "/\\jbhunt.com\\appshare\\FilesToBizlink\\Test-NonEDIFiles\\";
    /**
     * This function will be invoked when a new or updated blob is detected at the specified path. The blob contents are provided as input to this function.
     */
    @FunctionName("ExcelTOCSV")
    @StorageAccount("ordermanagementblob")
    public void blobHandler(
        @BlobTrigger(name = "content", path = "ordermanagementblob/omexceltocsv", dataType = "binary") byte[] content,
        @BindingName("name") String name,
        final ExecutionContext context
    ) {
        context.getLogger().info("Java Blob trigger function processed a blob. Name: " + name + "\n  Size: " + content.length + " Bytes");
        ExcelToCSVHanlder excelToCSVHanlder = new ExcelToCSVHanlder();
        excelToCSVHanlder.handleFile(content, new File(destination+"test.csv"));
    }
}
