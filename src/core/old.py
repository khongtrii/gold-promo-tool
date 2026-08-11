class Template_Mapping:

    CATEGORY_RULES: Optional[List] = [
        (
            re.compile(
                r"(?i)\b(front\s*page|back\s*page|unbeat)\b"
            ),
            "HERO",
        ),
        (
            re.compile(
                r"(?i)\b(cata|catalog(?:ue)?|fair|member\s*price|banner|exclusive\s*pack|family|other)\b"
            ),
            "CATA",
        ),
        (
            re.compile(
                r"(?i)\b(comple(?:mentary)?|comple)\b"
            ),
            "COMPLE",
        ),
        (
            re.compile(
                r"(?i)\bbuy\s*more\s*save\s*more\b"
            ),
            "STAR",
        ),
    ]
    
    def __init__(
        self,
        etl: Template_ETL
    ):

        self.src = etl.src
        self.nw = etl.dict_network
        self.plan = etl.plan
        self.cata = etl.cata
        self.cata_description = etl.cata_description
        self.cata_period = etl.cata_period

        self.allocation: dict = dict()

        self.template_check_oa: Optional[pd.DataFrame] = None
        self.template_promotion_plan: Optional[pd.DataFrame] = None
        self.template_update_so: Optional[pd.DataFrame] = None
        self.template_missing_ou: Optional[pd.DataFrame] = None
        self.template_so_calendar: Optional[pd.DataFrame] = None
        self.template_purchase: Optional[pd.DataFrame] = None
        self.template_po_commitment: Optional[pd.DataFrame] = None
        self.template_supplier_schedule: Optional[pd.DataFrame] = None
        self.template_add_attribute_marketing: Optional[pd.DataFrame] = None

    def attribute_map(self, text: str) -> str:
        if not text:
            return text
    
        text = str(text).strip()
    
        for pattern, value in self.CATEGORY_RULES:
            if pattern.search(text):
                return value
    
        return text
    
    @staticmethod
    def reset_no(data):
        data = data.reset_index(drop=True)
        data.insert(0, "NO", range(1, len(data) + 1))
        return data

    @staticmethod
    def get_contract(row):
        contract = row["CONTRACT"]
        site = row["SITE"]
        supplier = row["SUPPLIER"]
    
        if len(contract) == 8:
            return contract
        
        if supplier in etl.dict_network.get("wh8"):
            return f"{contract}0{supplier}"
    
        if len(site) == 3:
            return f"{contract}0{site}"
        elif len(site) == 4:
            return f"{contract}{site}"

    def fast_stage(self, data, have_no = False):
        data = data.drop_duplicates()
        data = data.dropna()
        data = data.reset_index(drop=True)
        if have_no:
            data = self.reset_no(data)
        return data
    
    def _create_check_oa(self) -> "Template_Mapping":
        data = self.src

        template_check_oa = {
            column_check_oa[0]:data["GOLD CODE"],
            column_check_oa[1]:data["LV"],
            column_check_oa[2]:data["LU"],
            column_check_oa[3]:data["PURCHASE NETWORK EXPANDED"],
            column_check_oa[4]:data["SUPPLIER CODE"],
            column_check_oa[5]:"1",
            column_check_oa[6]:data["COMMERCIAL CONTRACT"],
            column_check_oa[7]:data["PP START DATE"].dt.strftime("%d/%m/%Y"),
            column_check_oa[8]:data["PP END DATE"].dt.strftime("%d/%m/%Y")
        }

        template_check_oa = pd.DataFrame(template_check_oa)

        template_check_oa["SITE"] = template_check_oa["SITE"].str.split(";")

        template_check_oa = (
            template_check_oa
            .explode("SITE")
        )

        template_check_oa["CONTRACT"] = template_check_oa.apply(self.get_contract, axis=1)

        template_check_oa = template_check_oa[column_check_oa]

        template_check_oa = self.fast_stage(template_check_oa, have_no=True)

        self.template_check_oa = template_check_oa
        
        return self

    def _create_promotion_plan(self) -> "Template_Mapping":
        data = self.src
        
        template_promotion_plan = {
            column_promotion_plan[0]:data["SO"],
            column_promotion_plan[1]:f"{self.cata} {self.cata_description} ({self.cata_period})",
            column_promotion_plan[2]:f"{self.cata}D",
            column_promotion_plan[3]:data["SITE GROUP"],
            column_promotion_plan[4]:"1",
            column_promotion_plan[5]:self.plan["CATALOGUE START"].dt.strftime("%d/%m/%Y")[0],
            column_promotion_plan[6]:self.plan["CATALOGUE END"].dt.strftime("%d/%m/%Y")[0],
            column_promotion_plan[7]:self.plan["GLOBAL PERIOD START"][0],
            column_promotion_plan[8]:self.plan["GLOBAL PERIOD END"][0],
            column_promotion_plan[9]:self.plan["SHOP ACTIVATION"][0],
            column_promotion_plan[10]:self.plan["COMMITMENT DEADLINE"][0],
            column_promotion_plan[11]:self.plan["COMMITMENT CLOSING"][0],
            column_promotion_plan[12]:self.plan["ORDER WAREHOUSE START"][0],
            column_promotion_plan[13]:self.plan["ORDER WAREHOUSE END"][0]
        }

        template_promotion_plan = pd.DataFrame(template_promotion_plan)

        template_promotion_plan = template_promotion_plan[column_promotion_plan]
        
        template_promotion_plan = self.fast_stage(template_promotion_plan, have_no=False)
        
        self.template_promotion_plan = template_promotion_plan

        return self

    def _create_update_so(self) -> "Template_Mapping":
        data = self.src

        template_update_so = {
            column_update_so[0]:"1",
            column_update_so[1]:data["SO"],
            column_update_so[2]:data["GOLD CODE"],
            column_update_so[3]:data["LV"],
            column_update_so[4]:data["LU"]
        }

        template_update_so = pd.DataFrame(template_update_so)

        template_update_so = template_update_so[column_update_so]
        
        template_update_so = self.fast_stage(template_update_so, have_no=True)
        
        self.template_update_so = template_update_so

        return self

    def _create_missing_ou(self) -> "Template_Mapping":
        data = self.src

        template_missing_ou = {
            column_missing_ou[0]:data["SO"],
            column_missing_ou[1]:data["PURCHASE NETWORK EXPANDED"],
            column_missing_ou[2]:data["GOLD CODE"],
            column_missing_ou[3]:data["LV"],
            column_missing_ou[4]:data["LU"],
            column_missing_ou[5]:data["SUPPLIER CODE"],
            column_missing_ou[6]:"1",
            column_missing_ou[7]:data["COMMERCIAL CONTRACT"],
            column_missing_ou[8]:"1",
            column_missing_ou[9]:"10"
        }

        template_missing_ou = pd.DataFrame(template_missing_ou)

        template_missing_ou["SITE"] = template_missing_ou["SITE"].str.split(";")

        template_missing_ou = (
            template_missing_ou
            .explode("SITE")
        )

        template_missing_ou["CONTRACT"] = template_missing_ou.apply(self.get_contract, axis=1)

        template_missing_ou = template_missing_ou[column_missing_ou]
        
        template_missing_ou = self.fast_stage(template_missing_ou, have_no=True)
        
        self.template_missing_ou = template_missing_ou

        return self

    @staticmethod
    def _get_calendar(x):
        delivery_type = x["DELIVERY TYPE"]
    
        if delivery_type == "CROSS-DOCKING":
            order_cols = [
                "CROSS-DOCKING | ORDER DATE 1",
                "CROSS-DOCKING | ORDER DATE 2",
                "CROSS-DOCKING | ORDER DATE 3",
            ]
        elif delivery_type == "DIRECT":
            order_cols = [
                "DIRECT | ORDER DATE 1",
                "DIRECT | ORDER DATE 2",
                "DIRECT | ORDER DATE 3",
            ]
        else:
            order_cols = [
                "VINAMILK | ORDER DATE 1",
                "VINAMILK | ORDER DATE 2",
                "VINAMILK | ORDER DATE 3",
            ]
    
        delivery_cols = [
            "DELIVERY DATE 1",
            "DELIVERY DATE 2",
            "DELIVERY DATE 3",
        ]
    
        order_date_all = etl.plan[order_cols].iloc[0].tolist()
        delivery_date_all = etl.plan[delivery_cols].iloc[0].tolist()
    
        deli_values = x["%DELI"]
    
        order_date = []
        delivery_date = []
    
        for i, value in enumerate(deli_values):
            has_value = pd.notna(value) and str(value).strip() != ""
    
            if has_value:
                order_date.append(order_date_all[i])
                delivery_date.append(delivery_date_all[i])
    
        return pd.Series(
            {
                "ORDER DATE": order_date,
                "DELIVERY DATE": delivery_date,
            }
        )
    
    def _create_so_calendar(self) -> "Template_Mapping":
        data = self.src
        
        template_so_calendar = {
            column_so_calendar[0]:data["SO"],
            column_so_calendar[1]:data["PURCHASE NETWORK EXPANDED"],
            column_so_calendar[2]:data["SUPPLIER CODE"],
            column_so_calendar[3]:data["COMMERCIAL CONTRACT"],
            column_so_calendar[4]:"1",
            column_so_calendar[5]:data["GOLD CODE"],
            column_so_calendar[6]:data["LV"],
            column_so_calendar[7]:data["LU"],
            column_so_calendar[-3]:data[
                ["% DELIVERY 1", "% DELIVERY 2", "% DELIVERY 3"]
            ].values.tolist(),
            column_so_calendar[-2]:"10",
            column_so_calendar[-1]:"",
            "DELIVERY TYPE":data["DELIVERY TYPE"],
        }

        template_so_calendar = pd.DataFrame(template_so_calendar)
        
        template_so_calendar[["ORDER DATE", "DELIVERY DATE"]] = template_so_calendar.apply(self._get_calendar, axis=1)

        template_so_calendar.rename(columns={"%DELI":"PCT WEIGHT"}, inplace=True)
        
        template_so_calendar = template_so_calendar.explode([
            "PCT WEIGHT", "ORDER DATE", "DELIVERY DATE"
        ])
        
        template_so_calendar["SITE"] = template_so_calendar["SITE"].str.split(";")

        template_so_calendar = (
            template_so_calendar
            .explode("SITE")
        )

        template_so_calendar["CONTRACT"] = template_so_calendar.apply(self.get_contract, axis=1)

        template_so_calendar = template_so_calendar[column_so_calendar]

        template_so_calendar = self.fast_stage(template_so_calendar, have_no=True)

        self.template_so_calendar = template_so_calendar
        
        return self

    def _create_purchase(self) -> "Template_Mapping":
        data = self.src

        template_purchase = {
            column_purchase[0]:data["GOLD CODE"],
            column_purchase[1]:data["LV"],
            column_purchase[2]:"1",
            column_purchase[3]:data["NORMAL PURCHASE PRICE"],
            column_purchase[4]:data["PURCHASE NETWORK EXPANDED"],
            column_purchase[5]:data["PP START DATE"].dt.strftime("%d/%m/%Y"),
            column_purchase[6]:data["PP END DATE"].dt.strftime("%d/%m/%Y"),
            column_purchase[7]:data["COMMERCIAL CONTRACT"],
            column_purchase[8]:data["PURCHASE VAT"].map(VAT),
            column_purchase[9]:data["SUPPLIER CODE"],
            column_purchase[10]:"0"
        }
        
        template_purchase = pd.DataFrame(template_purchase)

        template_purchase["SITE"] = template_purchase["SITE"].str.split(";")

        template_purchase = (
            template_purchase
            .explode("SITE")
        )

        template_purchase["CONTRACT"] = template_purchase.apply(self.get_contract, axis=1)

        template_purchase = template_purchase[column_purchase]

        template_purchase = self.fast_stage(template_purchase, have_no=True)

        self.template_purchase = template_purchase
        
        return self

    def _get_allocation(self) -> "Template_Mapping":
        data = self.src

        site_column = [i for i in data.columns if i in self.nw["store"]]
        
        allocation = data[["GOLD CODE", "LV", "SUPPLIER CODE", "PURCHASE NETWORK EXPANDED", *site_column]]

        allocation["PURCHASE NETWORK EXPANDED"] = allocation["PURCHASE NETWORK EXPANDED"].str.split(";")

        allocation = (
            allocation
            .explode("PURCHASE NETWORK EXPANDED")
        )
        
        allocation["KEY"] = (
            allocation["GOLD CODE"].astype(str)
            + "-"
            + allocation["LV"].astype(str)
            + "-"
            + allocation["SUPPLIER CODE"].astype(str)
            + "-"
            + allocation["PURCHASE NETWORK EXPANDED"]
        )
        
        allocation["VALUE"] = allocation.apply(
            lambda x: x[x["PURCHASE NETWORK EXPANDED"]],
            axis=1,
        )
        
        network_dict = dict(zip(allocation["KEY"], allocation["VALUE"]))

        self.allocation = network_dict

        return self

    def _create_po_commitment(self) -> "Template_Mapping":
        self._get_allocation()
        
        data = self.src

        template_po_commitment = {
            column_po_commitment[0]:data["SO"],
            column_po_commitment[1]:data["SO"],
            column_po_commitment[2]:data["GOLD CODE"],
            column_po_commitment[3]:data["LV"],
            column_po_commitment[4]:data["LU"],
            column_po_commitment[5]:data["PURCHASE NETWORK EXPANDED"],
            column_po_commitment[7]:"2",
            column_po_commitment[8]:data["SUPPLIER CODE"]
        }

        template_po_commitment = pd.DataFrame(template_po_commitment)
        
        template_po_commitment["SITE"] = template_po_commitment["SITE"].str.split(";")
        
        template_po_commitment = template_po_commitment.explode("SITE")
        
        template_po_commitment["KEY"] = (
                    template_po_commitment["GOLD CODE"].astype(str)
                    + "-"
                    + template_po_commitment["LV"].astype(str)
                    + "-"
                    + template_po_commitment["SUPPLIER"].astype(str)
                    + "-"
                    + template_po_commitment["SITE"]
                )
        
        template_po_commitment = {
            **template_po_commitment,
            column_po_commitment[6]:template_po_commitment["KEY"].map(self.allocation)
        }
        
        template_po_commitment = pd.DataFrame(template_po_commitment)

        template_po_commitment = template_po_commitment[column_po_commitment]

        template_po_commitment = self.fast_stage(template_po_commitment, have_no=True)

        self.template_po_commitment = template_po_commitment
        
        return self

    def _create_supplier_schedule(self) -> "Template_Mapping":
        data = self.src
        
        template_supplier_schedule = {
            column_supplier_schedule[0]:data["SO"],
            column_supplier_schedule[1]:data["PURCHASE NETWORK EXPANDED"],
            column_supplier_schedule[2]:data["SUPPLIER CODE"],
            column_supplier_schedule[3]:"1",
            column_supplier_schedule[4]:data["COMMERCIAL CONTRACT"],
            column_supplier_schedule[5]:"",
            column_supplier_schedule[7]:"0000",
            column_supplier_schedule[9]:"2359",
            column_supplier_schedule[10]:"1",
            column_supplier_schedule[11]:"1",
            column_supplier_schedule[12]:"",
            "%DELI":data[
                ["% DELIVERY 1", "% DELIVERY 2", "% DELIVERY 3"]
            ].values.tolist(),
            "DELIVERY TYPE":data["DELIVERY TYPE"],
        }
        
        template_supplier_schedule = pd.DataFrame(template_supplier_schedule)
        
        template_supplier_schedule[["ORDER DATE", "DELIVERY DATE"]] = template_supplier_schedule.apply(self._get_calendar, axis=1)
        
        template_supplier_schedule = (
            template_supplier_schedule
            .explode(["%DELI", "ORDER DATE", "DELIVERY DATE"])
        )

        template_supplier_schedule["SITE"] = template_supplier_schedule["SITE"].str.split(";")

        template_supplier_schedule = (
            template_supplier_schedule
            .explode("SITE")
        )

        template_supplier_schedule["CONTRACT"] = template_supplier_schedule.apply(self.get_contract, axis=1)

        template_supplier_schedule = template_supplier_schedule[column_supplier_schedule]

        template_supplier_schedule = self.fast_stage(template_supplier_schedule, have_no=True)

        self.template_supplier_schedule = template_supplier_schedule
        
        return self

    def _create_add_attribute_marketing(self) -> "TemplateMapping":
        data = self.src
        
        template_add_attribute_marketing = {
            column_add_attribute_marketing[0]:"1",
            column_add_attribute_marketing[1]:data["SO"],
            column_add_attribute_marketing[2]:data["GOLD CODE"],
            column_add_attribute_marketing[3]:data["LV"],
            column_add_attribute_marketing[4]:data["LU"],
            column_add_attribute_marketing[5]:data["ATTRIBUTE MARKETING"],
            column_add_attribute_marketing[6]:data["FREE PRODUCT"]
        }

        template_add_attribute_marketing = pd.DataFrame(template_add_attribute_marketing)

        template_add_attribute_marketing["MEDIUM"] = template_add_attribute_marketing["MEDIUM"].map(self.attribute_map)

        template_add_attribute_marketing = template_add_attribute_marketing[column_add_attribute_marketing]

        template_add_attribute_marketing = self.fast_stage(template_add_attribute_marketing, have_no=True)

        self.template_add_attribute_marketing = template_add_attribute_marketing
        
        return self

    class Discount:
        def __init__(self, username:str = "user"):
            self.username = username

            self.template_ag: Optional[pd.DataFrame] = None

        def _create_ag(self):
            today = pd.Timestamp.today().strftime("%d.%m")
            
            template_ag = {
                column_ag[0]: "0",
                column_ag[1]: data["STRUCTURE"].str[:4],
                column_ag[2]: data["PURCHASE NETWORK EXPANDED"],
                column_ag[3]: data["SUPPLIER CODE"],
                column_ag[4]: data["COMMERCIAL CONTRACT"],
                column_ag[6]: "",
                column_ag[7]: f"{etl.cata}D.GP{etl.dept}-{self.username}({today})",
                column_ag[8]: data["PP START DATE"].min().strftime("%d/%m/%Y"),
                column_ag[9]: data["PP END DATE"].max().strftime("%d/%m/%Y"),
                column_ag[10]: data["GOLD CODE"],
                column_ag[11]: data["LV"],
                column_ag[12]: data["PP START DATE"].min().strftime("%d/%m/%Y"),
                column_ag[13]: data["PP END DATE"].max().strftime("%d/%m/%Y"),
                column_ag[14]: "0",
                column_ag[15]: ""
            }
            
            template_ag = pd.DataFrame(template_ag)
            
            template_ag[column_ag[2]] = template_ag[column_ag[2]].str.split(";")
            template_ag = (
                template_ag
                .explode(column_ag[2])
            )
            
            site_padded = template_ag[column_ag[2]].astype(str).str.zfill(4)
            supplier = template_ag[column_ag[3]].astype(str)
            
            key = supplier + site_padded
            sequence = key.groupby(key).cumcount() + 1
            sequence_str = sequence.apply(lambda x: f"{x:02d}")
            
            column_ag5_value = supplier + site_padded + sequence_str
            
            template_ag = {**template_ag, column_ag[5]: column_ag5_value}
            
            template_ag = pd.DataFrame(template_ag)

            template_ag["CONTRACT"] = template_ag.apply(self.get_contract, axis=1)
    
            template_ag = template_ag[column_ag]
    
            template_ag = self.fast_stage(template_ag, have_no=True)
    
            self.template_ag = template_ag
            
            return self