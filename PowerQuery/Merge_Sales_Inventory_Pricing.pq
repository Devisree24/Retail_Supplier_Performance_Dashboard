// =============================================================================
// Retail Supplier Performance - Power Query: Merge Sales + Inventory + Pricing
// In Excel: Data > Get Data > From File > From Text/CSV
// Load Sales, Inventory, Pricing_Promo as separate queries, then create Merged.
// =============================================================================
// Assumes queries are named: Sales, Inventory, Pricing_Promo

let
    // --- Load and clean Sales ---
    SalesRaw = Sales,
    SalesCleaned = Table.TransformColumnTypes(SalesRaw,{
        {"Date", type date},
        {"SKU", type text},
        {"Product Category", type text},
        {"Retailer", type text},
        {"Store", type text},
        {"Region", type text},
        {"Units Sold", Int64.Type},
        {"Revenue", type number},
        {"Cost", type number},
        {"Gross Margin %", type number}
    }),
    SalesDedup = Table.Distinct(SalesCleaned, {"Date", "SKU", "Retailer", "Store"}),
    SalesNoNulls = Table.ReplaceValue(SalesDedup, null, "", Replacer.ReplaceValue, {"SKU", "Retailer", "Store"}),
    SalesTrim = Table.TransformColumns(SalesNoNulls, {{"SKU", Text.Trim}, {"Retailer", Text.Trim}, {"Store", Text.Trim}}),

    // --- Load and clean Inventory ---
    InvRaw = Inventory,
    InvCleaned = Table.TransformColumnTypes(InvRaw,{
        {"SKU", type text},
        {"Retailer", type text},
        {"Store", type text},
        {"Region", type text},
        {"Beginning Inventory", Int64.Type},
        {"Ending Inventory", Int64.Type},
        {"Reorder Point", Int64.Type},
        {"Lead Time (days)", Int64.Type}
    }),
    InvNoNulls = Table.ReplaceValue(InvCleaned, null, 0, Replacer.ReplaceValue, {"Beginning Inventory", "Ending Inventory", "Reorder Point", "Lead Time (days)"}),
    InvTrim = Table.TransformColumns(InvNoNulls, {{"SKU", Text.Trim}, {"Retailer", Text.Trim}, {"Store", Text.Trim}}),

    // --- Load and clean Pricing & Promo (List Price, Promo Price, Promo Start, Promo End) ---
    PriceRaw = Pricing_Promo,
    PriceCleaned = Table.TransformColumnTypes(PriceRaw,{
        {"SKU", type text},
        {"Retailer", type text},
        {"List Price", type number},
        {"Promo Price", type number},
        {"Promo Start", type date},
        {"Promo End", type date},
        {"Marketing Spend", type number}
    }),
    PriceNoNulls = Table.ReplaceValue(PriceCleaned, null, 0, Replacer.ReplaceValue, {"Marketing Spend"}),
    PriceTrim = Table.TransformColumns(PriceNoNulls, {{"SKU", Text.Trim}, {"Retailer", Text.Trim}}),

    // --- Merge: Sales + Inventory (on SKU, Retailer, Store) ---
    MergedSI = Table.NestedJoin(SalesTrim, {"SKU", "Retailer", "Store"}, InvTrim, {"SKU", "Retailer", "Store"}, "Inventory", JoinKind.LeftOuter),
    ExpandedInv = Table.ExpandTableColumn(MergedSI, "Inventory",
        {"Region", "Beginning Inventory", "Ending Inventory", "Reorder Point", "Lead Time (days)"},
        {"Inv_Region", "Beginning Inventory", "Ending Inventory", "Reorder Point", "Lead Time (days)"}),

    // --- Merge with Pricing (on SKU, Retailer) ---
    MergedAll = Table.NestedJoin(ExpandedInv, {"SKU", "Retailer"}, PriceTrim, {"SKU", "Retailer"}, "Pricing", JoinKind.LeftOuter),
    ExpandedPrice = Table.ExpandTableColumn(MergedAll, "Pricing",
        {"List Price", "Promo Price", "Promo Start", "Promo End", "Marketing Spend"},
        {"List Price", "Promo Price", "Promo Start", "Promo End", "Marketing Spend"})
in
    ExpandedPrice
