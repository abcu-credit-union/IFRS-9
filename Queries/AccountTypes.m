let
    EffDate = Date.FromText(extractPers{0}[Branch Province]),
    BCUSource = Oracle.Database("BCUDatabase", [Query=
        "SELECT
            'BCU' AS SOURCE,
            ACCTNBR,
            MJACCTTYPCD,
            CURRMIACCTTYPCD,
            CURRACCTSTATCD,
            TAXRPTFORPERSNBR
        FROM ACCT
        WHERE CURRACCTSTATCD <> 'CLS'
            AND MJACCTTYPCD <> 'GL'
    "]),
    RCCUSource = Oracle.Database("RCCUDatabase", [Query=
        "SELECT
            'RCCU' AS SOURCE,
            ACCTNBR,
            MJACCTTYPCD,
            CURRMIACCTTYPCD,
            CURRACCTSTATCD,
            TAXRPTFORPERSNBR
        FROM ACCT
        WHERE CURRACCTSTATCD <> 'CLS'
            AND MJACCTTYPCD <> 'GL'
    "]),
    ABCUSource = Table.Combine({BCUSource, RCCUSource}),
    #"Changed Type" = Table.TransformColumnTypes(ABCUSource,{{"ACCTNBR", type text}}),
    #"Added Account Types" = Table.AddColumn(#"Changed Type", "Type", each fnAccountType([CURRMIACCTTYPCD]))
in
    #"Added Account Types"