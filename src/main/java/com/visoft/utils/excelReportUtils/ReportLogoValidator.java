package com.visoft.utils.excelReportUtils;

import com.visoft.dto.ReportLogo;

import java.util.Optional;

import static com.visoft.services.Extensions.imageExtensions;
import static com.visoft.utils.MineTypeValidator.getMineTypeFromByteArray;

public class ReportLogoValidator {

    public static boolean isValidReportLogo(ReportLogo reportLogo){
        Optional<ReportLogo> reportLogoOpt = Optional.ofNullable(reportLogo);
        if(reportLogoOpt.isPresent()){
            return reportLogoOpt
                    .filter(x->x.getData()!=null)
                    .filter(x->checkImageMineType(x.getData()))
                    .isPresent();
        }
        return false;
    }

    private static boolean checkImageMineType(byte[] data){
        boolean result = true;
        if(data!=null) try {
            result = imageExtensions.contains(getMineTypeFromByteArray(data));
        } catch (Throwable e ) {
            e.printStackTrace();
            result = false;
        }
        return result;
    }
}