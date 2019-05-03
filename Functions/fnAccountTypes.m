let
    fnAccountType = (CURRMIACCTTYPCD) =>
        let
            loan = "Loan",
            aod = "AOD",
            loc = "LOC",
            heloc = "HELOC",
            values = {
                {"LTCV", loan},
                {"LLCV", loc},
                {"CI04", aod},
                {"CP02", aod},
                {"CN99", heloc},
                {"LOCM", heloc},
                {"CP01", aod},
                {"CP04", aod},
                {"CI08", aod},
                {"CP05", aod},
                {"LTCR", loan},
                {"LTCF", loan},
                {"CI03", aod},
                {"CN19", aod},
                {"CP03", aod},
                {"CI12", aod},
                {"LLCR", loc},
                {"LLCS", loc},
                {"LLCF", loc},
                {"CI20", aod},
                {"CI02", aod},
                {"CI13", aod},
                {"CI11", aod},
                {"CN05", aod},
                {"CN09", aod},
                {"CI21", aod},
                {"CI19", aod},
                {"LTCX", loan}
                },
            accountType = 
                try
                    List.First(List.Select(values, each _{0}=CURRMIACCTTYPCD)){1}
                otherwise "Error - " & CURRMIACCTTYPCD & " has not been mapped"
        in
            accountType
in
    fnAccountType