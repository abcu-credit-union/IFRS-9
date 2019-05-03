let
    BCUSource = Oracle.Database("BCUDatabase", [Query=
        "SELECT
            'BCU' AS SOURCE,
            ACCTNBR,
            NOTEOPENAMT,
            OWNERNAME,
            TAXRPTFORPERSNBR
        FROM WH_ACCTCOMMON
        WHERE EFFDATE = TRUNC(CURRENT_DATE - 1)
	    AND CURRACCTSTATCD <> 'CLS'"]),
    RCCUSource = Oracle.Database("RCCUDatabase", [Query=
        "SELECT
            'RCCU' AS SOURCE,
            ACCTNBR,
            NOTEOPENAMT,
            OWNERNAME,
            TAXRPTFORPERSNBR
        FROM WH_ACCTCOMMON
        WHERE EFFDATE = TRUNC(CURRENT_DATE - 1)
	    AND CURRACCTSTATCD <> 'CLS'"]),
    ABCUSource = Table.Combine({BCUSource, RCCUSource}),
    #"Changed Type" = Table.TransformColumnTypes(ABCUSource,{{"ACCTNBR", type text}})
in
    #"Changed Type"