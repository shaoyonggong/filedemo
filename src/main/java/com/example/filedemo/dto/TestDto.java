package com.example.filedemo.dto;

import lombok.*;
import lombok.experimental.Accessors;

import java.io.Serializable;
import java.math.BigDecimal;

/**
 * @Description
 * @Author shaoyonggong
 * @Date 2020/2/28
 */
@Data
@AllArgsConstructor
@NoArgsConstructor
@Accessors(chain = true)
@EqualsAndHashCode(callSuper = false)
@Builder
public class TestDto implements Serializable {

    private static final long serialVersionUID = 1L;

    private String sourceCode;

    private String po;

    private String business;

    private String platform;

    private String stockOrgNo;

    private String stockOrgName;

    private String logicWarehouseNo;

    private String logicWarehouseName;

    private BigDecimal num;

}
