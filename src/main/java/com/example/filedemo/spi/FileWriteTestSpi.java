package com.example.filedemo.spi;

import com.example.filedemo.dto.TestDto;
import com.example.filedemo.utils.ExportUtil;
import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.math.BigDecimal;
import java.net.URLEncoder;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

/**
 * @Description
 * @Author shaoyonggong
 * @Date 2020/4/1
 */
@Api(tags = {"测试生成excel", "分类:文件"})
@RequestMapping(value = "/generate/", method = RequestMethod.GET)
@RestController
public class FileWriteTestSpi {
    //excel标题
    private static List<String> title = Arrays.asList("id", "sourceCode", "po", "business", "platform", "stockOrgNo", "stockOrgName", "logicWarehouseNo", "logicWarehouseName", "totalNum");

    @ResponseBody
    @ApiOperation("jjjj")
    @RequestMapping
    public List<String> download(HttpServletRequest request, HttpServletResponse response) throws IOException, ServletException {
        List<String> returnList = new ArrayList<>();
        // 数据
        List<TestDto> list = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            list.add(TestDto.builder()
                    .sourceCode("sourceCode" + i)
                    .po("po" + i)
                    .business("business" + i)
                    .platform("platform" + 1)
                    .stockOrgNo("stockOrgNo" + i)
                    .stockOrgName("stockOrgName" + i)
                    .logicWarehouseNo("logicWarehouseNo" + i)
                    .logicWarehouseName("logicWarehouseName" + i)
                    .num(getRandomRedPacketBetweenMinAndMax(new BigDecimal(0.0000), new BigDecimal(1000.0000)))
                    .build());
        }
        sendExcel(list, response);
        returnList.add("chenggong");
        return returnList;
    }

    public void sendExcel(List<TestDto> list, HttpServletResponse response) throws IOException, ServletException {

        List<List> contentList = new ArrayList<>();

        for (int i = 0; i < list.size(); i++) {
            TestDto f = list.get(i);
            List rowContent = new ArrayList<>();
            rowContent.add(String.valueOf(i));
            rowContent.add(String.valueOf(f.getSourceCode()));
            rowContent.add(String.valueOf(f.getPo()));
            rowContent.add(String.valueOf(f.getBusiness()));
            rowContent.add(String.valueOf(f.getPlatform()));
            rowContent.add(String.valueOf(f.getStockOrgNo()));
            rowContent.add(String.valueOf(f.getStockOrgName()));
            rowContent.add(String.valueOf(f.getLogicWarehouseNo()));
            rowContent.add(String.valueOf(f.getLogicWarehouseName()));
            rowContent.add(String.valueOf(f.getNum().setScale(4, BigDecimal.ROUND_UP)));
            contentList.add(rowContent);
        }
        //sheet名
        String sheetName = "测试文件sheet1";
        //创建HSSFWorkbook
        XSSFWorkbook wb = ExportUtil.getXSSFWorkbook(sheetName, title, contentList);

        try {
            //文件路径：  SaleAppointmentSendEmail/日期/
            String datePath = new SimpleDateFormat("yyyyMMdd").format(new Date());
            String filePath = "SaleAppointmentSendEmail/" + datePath + "/";
            String path = "/exportexcel/test/" + filePath;

            isChartPathExist(path);
            //文件名：仓库/承运商 + 日期.xlsx
            FileOutputStream fos = new FileOutputStream(path + "key-" + datePath + ".xlsx");

            wb.write(fos);
            fos.close();
            wb.close();
        } catch (IOException e) {
            e.printStackTrace();
            wb.close();
        }

//        //doGet(response);
//        //获取文件
//        try{
//            jxl.Workbook excelWb =null;
//            InputStream is = new FileInputStream("/test/SaleAppointmentSendEmail/20200407/仓库、承运商20200407.xlsx");
//            excelWb = Workbook.getWorkbook(is);
//
//            int sheetSize = excelWb.getNumberOfSheets();
//            Sheet sheet = excelWb.getSheet(0);
//            int row_total = sheet.getRows();
//            for (int j = 0; j < row_total; j++) {
//                if(j == 0){
//                    Cell[] cells = sheet.getRow(j);
//
//                    System.out.println(cells[0].getContents());
//                    System.out.println(cells[1].getContents());
//                    System.out.println(cells[2].getContents());
//                }
//            }
//        }catch (IOException e) {
//            // TODO Auto-generated catch block
//            e.printStackTrace();
//        } catch (BiffException e){
//            e.printStackTrace();
//        }

    }

    /**
     * 获取金额
     *
     * @param min
     * @param max
     * @return
     */
    public static BigDecimal getRandomRedPacketBetweenMinAndMax(BigDecimal min, BigDecimal max) {
        float minF = min.floatValue();
        float maxF = max.floatValue();

        //生成随机数
        BigDecimal db = new BigDecimal(Math.random() * (maxF - minF) + minF);

        //返回保留两位小数的随机数。不进行四舍五入
        return db.setScale(2, BigDecimal.ROUND_DOWN);
    }

    /**
     * 判断文件夹是否存在不存在则创建
     *
     * @param dirPath
     */
    private static void isChartPathExist(String dirPath) {

        File file = new File(dirPath);
        if (!file.exists()) {
            file.mkdirs();
        }
    }

    private void doGet(HttpServletResponse resp)
            throws ServletException, IOException {

        String path = "/test/SaleAppointmentSendEmail/20200407/仓库、承运商20200407.xlsx";

        resp.setContentType("application/x-download");

        resp.setHeader("content-disposition", "attachment;filename="
                + URLEncoder.encode("仓库、承运商20200407.xlsx", "UTF-8"));

        InputStream input = null;
        OutputStream output = null;
        try {
            input = new FileInputStream(path);
            System.out.println(input.available());
            output = resp.getOutputStream();
            int len = 0;
            byte bts[] = new byte[1024];
            while ((len = input.read(bts)) != -1) {
                output.write(bts, 0, len);
            }
        } finally {
            if (input != null) {
                input.close();
            }
        }
    }
}
