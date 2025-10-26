package com.labysolutions.j_hybridqc;


import lombok.AllArgsConstructor;
import lombok.NoArgsConstructor;

//@AllArgsConstructor
//@NoArgsConstructor(force = true)
public class HybridLogic {

    private final String filename;
    private final String saveas;
    private final Integer min_missing_percentage;
    private final Integer min_perc_polymorphic;
    private final Integer min_perc_hybridity;


    private final String green = "00FF00";
    private final String blue = "007BFF";
    private final String red = "FF0000";
    private final String orange = "FFA500";
    private final String pink = "FF00FF";
    private final String grey = "C0C0C0";

    public HybridLogic(
            String filename,
            String saveas, Integer minMissingPercentage,
            Integer minPercPolymorphic,
            Integer min_perc_hybridity
    ){
        this.filename = filename;
        this.saveas = saveas;
        this.min_missing_percentage = minMissingPercentage;
        this.min_perc_polymorphic = minPercPolymorphic;
        this.min_perc_hybridity = min_perc_hybridity;
    }


    String test_filename = "data.xlsx";



}
