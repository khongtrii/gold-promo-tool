import re
from typing import List, Optional

import pandas as pd

from src.constant.required import VAT
from src.constant.template import (
    column_add_attribute_marketing,
    column_ag,
    column_attr,
    column_check_oa,
    column_dc,
    column_de,
    column_missing_ou,
    column_po_commitment,
    column_promotion_plan,
    column_purchase,
    column_sale,
    column_so_calendar,
    column_supplier_schedule,
    column_update_so,
    report,
)
from src.service.template_service import Template_ETL


class ContractMixin:
    """Logic ghép mã CONTRACT từ CONTRACT/SITE/SUPPLIER."""

    def get_contract(self, row):
        contract = "" if pd.isna(row["CONTRACT"]) else str(row["CONTRACT"]).strip()
        site = "" if pd.isna(row["SITE"]) else str(row["SITE"]).strip()
        supplier = "" if pd.isna(row["SUPPLIER"]) else str(row["SUPPLIER"]).strip()

        if len(contract) == 8:
            return contract

        if supplier in self.nw.get("wh8"):
            return f"{contract}0{supplier}"

        if len(site) == 3:
            return f"{contract}0{site}"
        elif len(site) == 4:
            return f"{contract}{site}"
        return contract


class StageMixin:
    """Các bước xử lý DataFrame dùng chung cho mọi template."""

    @staticmethod
    def reset_no(data):
        data = data.reset_index(drop=True)
        data.insert(0, "NO", range(1, len(data) + 1))
        return data

    def fast_stage(self, data, have_no=False):
        data = data.drop_duplicates()
        # data = data.dropna()
        data = data.reset_index(drop=True)
        if have_no:
            data = self.reset_no(data)
        return data


class CalendarMixin:
    """Logic tính ORDER DATE / DELIVERY DATE theo DELIVERY TYPE."""

    def _get_calendar(self, x):
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

        order_date_all = self.plan[order_cols].iloc[0].tolist()
        delivery_date_all = self.plan[delivery_cols].iloc[0].tolist()

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


class AllocationMixin:
    """Xây self.allocation: map KEY -> VALUE để tra commitment quantity."""

    def _get_allocation(self) -> "AllocationMixin":
        data = self.src

        site_column = [i for i in data.columns if i in self.nw["store"]]

        allocation = data[
            ["GOLD CODE", "LV", "SUPPLIER CODE", "PURCHASE NETWORK EXPANDED", *site_column]
        ]

        allocation["PURCHASE NETWORK EXPANDED"] = allocation[
            "PURCHASE NETWORK EXPANDED"
        ].str.split(";")

        allocation = allocation.explode("PURCHASE NETWORK EXPANDED")

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


class AttributeMapMixin:
    """Map MEDIUM (free text) -> category chuẩn hoá bằng regex."""

    CATEGORY_RULES: Optional[List] = [
        (
            re.compile(r"(?i)\b(hero|front\s*page|back\s*page|unbeat)\b"),
            "HERO",
        ),
        (
            re.compile(
                r"(?i)\b(cata|catalog(?:ue)?|fair|member\s*price|banner|exclusive\s*pack|family|other|normal|the\s*1)\b"
            ),
            "CATA",
        ),
        (
            re.compile(r"(?i)\b(comple(?:mentary)?|comple)\b"),
            "COMPLE",
        ),
        (
            re.compile(r"(?i)\bbuy\s*more\s*save\s*more\b"),
            "STAR",
        ),
    ]

    def attribute_map(self, text: str) -> str:
        if not text:
            return text

        text = str(text).strip()

        for pattern, value in self.CATEGORY_RULES:
            if pattern.search(text):
                return value

        return text


# --------------------------------------------------------------------------- #
# Base: thuộc tính chung + __init__
# --------------------------------------------------------------------------- #
class BaseTemplate(ContractMixin, StageMixin):
    """
    Chứa __init__ và toàn bộ state dùng chung. Mọi mixin template khác đều
    giả định đã có sẵn: self.src, self.nw, self.plan, self.cata,
    self.cata_description, self.cata_period.
    """

    def __init__(self, etl: "Template_ETL"):
        self.src = etl.src
        self.nw = etl.dict_network
        self.plan = etl.plan
        self.cata = etl.cata
        self.cata_description = etl.cata_description
        self.cata_period = etl.cata_period

        self.allocation: dict = dict()

        self.template_check_oa: Optional[pd.DataFrame] = None
        self.template_check_oa_raw: Optional[pd.DataFrame] = None
        self.template_promotion_plan: Optional[pd.DataFrame] = None
        self.template_update_so: Optional[pd.DataFrame] = None
        self.template_missing_ou: Optional[pd.DataFrame] = None
        self.template_so_calendar: Optional[pd.DataFrame] = None
        self.template_purchase: Optional[pd.DataFrame] = None
        self.template_po_commitment: Optional[pd.DataFrame] = None
        self.template_supplier_schedule: Optional[pd.DataFrame] = None
        self.template_add_attribute_marketing: Optional[pd.DataFrame] = None
        
class CheckOAMixin:
    def _create_check_oa(self) -> "CheckOAMixin":
        data = self.src

        template_check_oa = {
            "STRUCTURE" : data['STRUCTURE'],
            "RAW NETWORK": data["GOLD PROMO NETWORK EXPANDED"],
            column_check_oa[0]: data["GOLD CODE"],
            column_check_oa[1]: data["LV"],
            column_check_oa[2]: data["LU"],
            column_check_oa[3]: data["PURCHASE NETWORK EXPANDED"],
            column_check_oa[4]: data["SUPPLIER CODE"],
            column_check_oa[5]: "1",
            column_check_oa[6]: data["COMMERCIAL CONTRACT"],
            column_check_oa[7]: data["PP START DATE"].dt.strftime("%d/%m/%Y"),
            column_check_oa[8]: data["PP END DATE"].dt.strftime("%d/%m/%Y"),
        }

        template_check_oa = pd.DataFrame(template_check_oa)

        template_check_oa["SITE"] = template_check_oa["SITE"].str.split(";")

        template_check_oa = template_check_oa.explode("SITE")

        template_check_oa["CONTRACT"] = template_check_oa.apply(self.get_contract, axis=1)

        self.template_check_oa_raw = template_check_oa.copy()

        template_check_oa = template_check_oa[column_check_oa]

        template_check_oa = self.fast_stage(template_check_oa, have_no=True)

        self.template_check_oa = template_check_oa

        return self


class PromotionPlanMixin:
    def _create_promotion_plan(self) -> "PromotionPlanMixin":
        data = self.src

        template_promotion_plan = {
            column_promotion_plan[0]: data["SO"],
            column_promotion_plan[1]: f"{self.cata} {self.cata_description} ({self.cata_period})",
            column_promotion_plan[2]: f"{self.cata}D",
            column_promotion_plan[3]: data["SITE GROUP"],
            column_promotion_plan[4]: "1",
            # ``_load_plan`` normalizes these source date columns to the
            # output date format before the mapping layer is invoked.
            column_promotion_plan[5]: self.plan["CATALOGUE START DATE"].iloc[0],
            column_promotion_plan[6]: self.plan["CATALOGUE END DATE"].iloc[0],
            column_promotion_plan[7]: self.plan["GLOBAL PERIOD START"].iloc[0],
            column_promotion_plan[8]: self.plan["GLOBAL PERIOD END"].iloc[0],
            column_promotion_plan[9]: self.plan["SHOP ACTIVATION"].iloc[0],
            column_promotion_plan[10]: self.plan["COMMITMENT DEADLINE"].iloc[0],
            column_promotion_plan[11]: self.plan["COMMITMENT CLOSING"].iloc[0],
            column_promotion_plan[12]: self.plan["ORDER WAREHOUSE START"].iloc[0],
            column_promotion_plan[13]: self.plan["ORDER WAREHOUSE END"].iloc[0],
        }

        template_promotion_plan = pd.DataFrame(template_promotion_plan)

        template_promotion_plan = template_promotion_plan[column_promotion_plan]

        template_promotion_plan = self.fast_stage(template_promotion_plan, have_no=False)

        self.template_promotion_plan = template_promotion_plan

        return self


class UpdateSOMixin:
    def _create_update_so(self) -> "UpdateSOMixin":
        data = self.src

        template_update_so = {
            column_update_so[0]: "1",
            column_update_so[1]: data["SO"],
            column_update_so[2]: data["GOLD CODE"],
            column_update_so[3]: data["LV"],
            column_update_so[4]: data["LU"],
        }

        template_update_so = pd.DataFrame(template_update_so)

        template_update_so = template_update_so[column_update_so]

        template_update_so = self.fast_stage(template_update_so, have_no=True)

        self.template_update_so = template_update_so

        return self


class MissingOUMixin:
    def _create_missing_ou(self) -> "MissingOUMixin":
        data = self.src

        template_missing_ou = {
            column_missing_ou[0]: data["SO"],
            column_missing_ou[1]: data["PURCHASE NETWORK EXPANDED"],
            column_missing_ou[2]: data["GOLD CODE"],
            column_missing_ou[3]: data["LV"],
            column_missing_ou[4]: data["LU"],
            column_missing_ou[5]: data["SUPPLIER CODE"],
            column_missing_ou[6]: "1",
            column_missing_ou[7]: data["COMMERCIAL CONTRACT"],
            column_missing_ou[8]: "1",
            column_missing_ou[9]: "10",
        }

        template_missing_ou = pd.DataFrame(template_missing_ou)

        template_missing_ou["SITE"] = template_missing_ou["SITE"].str.split(";")

        template_missing_ou = template_missing_ou.explode("SITE")

        template_missing_ou["CONTRACT"] = template_missing_ou.apply(self.get_contract, axis=1)

        template_missing_ou = template_missing_ou[column_missing_ou]

        template_missing_ou = self.fast_stage(template_missing_ou, have_no=True)

        self.template_missing_ou = template_missing_ou

        return self


class SOCalendarMixin(CalendarMixin):
    def _create_so_calendar(self) -> "SOCalendarMixin":
        data = self.src

        template_so_calendar = {
            column_so_calendar[0]: data["SO"],
            column_so_calendar[1]: data["PURCHASE NETWORK EXPANDED"],
            column_so_calendar[2]: data["SUPPLIER CODE"],
            column_so_calendar[3]: data["COMMERCIAL CONTRACT"],
            column_so_calendar[4]: "1",
            column_so_calendar[5]: data["GOLD CODE"],
            column_so_calendar[6]: data["LV"],
            column_so_calendar[7]: data["LU"],
            "%DELI": data[
                ["% DELIVERY 1", "% DELIVERY 2", "% DELIVERY 3"]
            ].values.tolist(),
            column_so_calendar[-2]: "10",
            column_so_calendar[-1]: "",
            "DELIVERY TYPE": data["DELIVERY TYPE"],
        }

        template_so_calendar = pd.DataFrame(template_so_calendar)

        template_so_calendar[["ORDER DATE", "DELIVERY DATE"]] = template_so_calendar.apply(
            self._get_calendar, axis=1
        )

        template_so_calendar.rename(columns={"%DELI": "PCT WEIGHT"}, inplace=True)

        template_so_calendar = template_so_calendar.explode(
            ["PCT WEIGHT", "ORDER DATE", "DELIVERY DATE"]
        )

        template_so_calendar["SITE"] = template_so_calendar["SITE"].str.split(";")

        template_so_calendar = template_so_calendar.explode("SITE")

        template_so_calendar["CONTRACT"] = template_so_calendar.apply(self.get_contract, axis=1)

        template_so_calendar = template_so_calendar[column_so_calendar]

        template_so_calendar = self.fast_stage(template_so_calendar, have_no=True)

        self.template_so_calendar = template_so_calendar

        return self


class PurchaseMixin:
    def _create_purchase(self) -> "PurchaseMixin":
        data = self.src

        template_purchase = {
            column_purchase[0]: data["GOLD CODE"],
            column_purchase[1]: data["LV"],
            column_purchase[2]: "1",
            column_purchase[3]: data["NORMAL PURCHASE PRICE"],
            column_purchase[4]: data["PURCHASE NETWORK EXPANDED"],
            column_purchase[5]: data["PP START DATE"].dt.strftime("%d/%m/%Y"),
            column_purchase[6]: data["PP END DATE"].dt.strftime("%d/%m/%Y"),
            column_purchase[7]: data["COMMERCIAL CONTRACT"],
            column_purchase[8]: data["PURCHASE VAT"].map(VAT).fillna(""),
            column_purchase[9]: data["SUPPLIER CODE"],
            column_purchase[10]: "0",
        }

        template_purchase = pd.DataFrame(template_purchase)

        template_purchase["SITE"] = template_purchase["SITE"].str.split(";")

        template_purchase = template_purchase.explode("SITE")

        template_purchase["CONTRACT"] = template_purchase.apply(self.get_contract, axis=1)

        template_purchase = template_purchase[column_purchase]

        template_purchase = self.fast_stage(template_purchase, have_no=True)

        self.template_purchase = template_purchase

        return self


class POCommitmentMixin(AllocationMixin):
    """Cần allocation (self.allocation) nên kế thừa AllocationMixin."""

    def _create_po_commitment(self) -> "POCommitmentMixin":
        self._get_allocation()

        data = self.src

        template_po_commitment = {
            column_po_commitment[0]: data["SO"],
            column_po_commitment[1]: data["SO"],
            column_po_commitment[2]: data["GOLD CODE"],
            column_po_commitment[3]: data["LV"],
            column_po_commitment[4]: data["LU"],
            column_po_commitment[5]: data["PURCHASE NETWORK EXPANDED"],
            column_po_commitment[7]: "2",
            "SUPPLIER": data["SUPPLIER CODE"],
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
            column_po_commitment[6]: template_po_commitment["KEY"].map(self.allocation),
        }

        template_po_commitment = pd.DataFrame(template_po_commitment)

        template_po_commitment = template_po_commitment[column_po_commitment]

        template_po_commitment = self.fast_stage(template_po_commitment, have_no=True)

        self.template_po_commitment = template_po_commitment

        return self


class SupplierScheduleMixin(CalendarMixin):
    def _create_supplier_schedule(self) -> "SupplierScheduleMixin":
        data = self.src

        template_supplier_schedule = {
            column_supplier_schedule[0]: data["SO"],
            column_supplier_schedule[1]: data["PURCHASE NETWORK EXPANDED"],
            column_supplier_schedule[2]: data["SUPPLIER CODE"],
            column_supplier_schedule[3]: "1",
            column_supplier_schedule[4]: data["COMMERCIAL CONTRACT"],
            column_supplier_schedule[5]: "",
            column_supplier_schedule[7]: "0000",
            column_supplier_schedule[9]: "2359",
            column_supplier_schedule[10]: "1",
            column_supplier_schedule[11]: "1",
            column_supplier_schedule[12]: "",
            "%DELI": data[["% DELIVERY 1", "% DELIVERY 2", "% DELIVERY 3"]].values.tolist(),
            "DELIVERY TYPE": data["DELIVERY TYPE"],
        }

        template_supplier_schedule = pd.DataFrame(template_supplier_schedule)

        template_supplier_schedule[["ORDER DATE", "DELIVERY DATE"]] = (
            template_supplier_schedule.apply(self._get_calendar, axis=1)
        )

        template_supplier_schedule = template_supplier_schedule.explode(
            ["%DELI", "ORDER DATE", "DELIVERY DATE"]
        )

        template_supplier_schedule["SITE"] = template_supplier_schedule["SITE"].str.split(";")

        template_supplier_schedule = template_supplier_schedule.explode("SITE")

        template_supplier_schedule["CONTRACT"] = template_supplier_schedule.apply(
            self.get_contract, axis=1
        )

        template_supplier_schedule = template_supplier_schedule[column_supplier_schedule]

        template_supplier_schedule = self.fast_stage(template_supplier_schedule, have_no=True)

        self.template_supplier_schedule = template_supplier_schedule

        return self


class AddAttributeMarketingMixin(AttributeMapMixin):
    def _create_add_attribute_marketing(self) -> "AddAttributeMarketingMixin":
        data = self.src

        template_add_attribute_marketing = {
            column_add_attribute_marketing[0]: "1",
            column_add_attribute_marketing[1]: data["SO"],
            column_add_attribute_marketing[2]: data["GOLD CODE"],
            column_add_attribute_marketing[3]: data["LV"],
            column_add_attribute_marketing[4]: data["LU"],
            column_add_attribute_marketing[5]: data["ATTRIBUTE MARKETING"],
            column_add_attribute_marketing[6]: data["FREE PRODUCT"],
        }

        template_add_attribute_marketing = pd.DataFrame(template_add_attribute_marketing)

        template_add_attribute_marketing["MEDIUM"] = template_add_attribute_marketing[
            "MEDIUM"
        ].map(self.attribute_map)

        template_add_attribute_marketing = template_add_attribute_marketing[
            column_add_attribute_marketing
        ]

        template_add_attribute_marketing = self.fast_stage(
            template_add_attribute_marketing, have_no=True
        )

        self.template_add_attribute_marketing = template_add_attribute_marketing

        return self

class Template_Mapping(
    BaseTemplate,
    CheckOAMixin,
    PromotionPlanMixin,
    UpdateSOMixin,
    MissingOUMixin,
    SOCalendarMixin,
    PurchaseMixin,
    POCommitmentMixin,
    SupplierScheduleMixin,
    AddAttributeMarketingMixin,
):
    pass


class DiscountTypeMixin:
    _NUMBER_PATTERN = re.compile(r"\d+(?:\.\d+)?")
 
    @staticmethod
    def get_discount_type(value) -> Optional[str]:
        if pd.isna(value):
            return None
 
        text = str(value)
 
        if "+" in text:
            return "3"
        if "%" in text:
            return "1"
        return "2"
 
    @classmethod
    def is_zero_discount(cls, value) -> bool:
        if pd.isna(value):
            return False
 
        text = str(value).strip()
        if text == "":
            return False
 
        numbers = cls._NUMBER_PATTERN.findall(text)
        if not numbers:
            return False
 
        return all(float(n) == 0 for n in numbers)
 
class Discount(ContractMixin, StageMixin, DiscountTypeMixin):
    def __init__(self, etl: "Template_ETL", username: str = "user"):
        self.src = etl.src
        self.etl = etl
        self.nw = etl.dict_network
        self.username = username
 
        self.template_ag_raw: Optional[pd.DataFrame] = None

        self.template_ag: Optional[pd.DataFrame] = None

        self.report_err: Optional[pd.DataFrame] = None
        
        self.template_dc_free: Optional[pd.DataFrame] = None
        self.template_dc_money: Optional[pd.DataFrame] = None
        
        self.template_de: Optional[pd.DataFrame] = None

    def _require_ag_raw(self) -> pd.DataFrame:
        if self.template_ag_raw is None:
            raise ValueError("Create template_ag before importing the AG report or creating discount templates.")
        return self.template_ag_raw
 
    def _create_ag_raw(self) -> "Discount":
        data = self.src
        etl = self.etl

        if data is None:
            raise ValueError("Stage 1 source data is not available for Discount processing.")
        if data.empty:
            self.template_ag_raw = pd.DataFrame(columns=[*column_ag, "DISCOUNT VALUE", "RAW START DATE", "RAW END DATE", "DISCOUNT TYPE", "CONTRACT"])
            return self

        start_date = data["PP START DATE"].min()
        end_date = data["PP END DATE"].max()
        if pd.isna(start_date) or pd.isna(end_date):
            raise ValueError("Discount processing requires valid PP START DATE and PP END DATE values.")

        today = pd.Timestamp.today().strftime("%d.%m")
 
        template_ag_raw = {
            column_ag[0]: "0",
            column_ag[1]: "0" + data["STRUCTURE"].astype(str) + "0",
            column_ag[2]: data["PURCHASE NETWORK EXPANDED"],
            column_ag[3]: data["SUPPLIER CODE"],
            column_ag[4]: data["COMMERCIAL CONTRACT"],
            column_ag[6]: "",
            column_ag[7]: (
                f"{etl.cata}D.GP"
                + data["STRUCTURE"].astype(str)
                + f"-{self.username}({today})"
            ),
            column_ag[8]: start_date.strftime("%d/%m/%Y"),
            column_ag[9]: end_date.strftime("%d/%m/%Y"),
            column_ag[10]: data["GOLD CODE"],
            column_ag[11]: data["LV"],
            column_ag[12]: start_date.strftime("%d/%m/%Y"),
            column_ag[13]: end_date.strftime("%d/%m/%Y"),
            column_ag[14]: "0",
            column_ag[15]: "",
            "DISCOUNT VALUE": data["DISCOUNT (% OR VALUE)"],
            "RAW START DATE": data["PP START DATE"].dt.strftime("%d/%m/%Y"),
            "RAW END DATE": data["PP END DATE"].dt.strftime("%d/%m/%Y"),
        }
 
        template_ag_raw = pd.DataFrame(template_ag_raw)
 
        template_ag_raw["DISCOUNT TYPE"] = template_ag_raw["DISCOUNT VALUE"].apply(
            self.get_discount_type
        )
        
        template_ag_raw = template_ag_raw[
            ~template_ag_raw["DISCOUNT VALUE"].apply(self.is_zero_discount)
        ].reset_index(drop=True)
 
        # template_ag_raw[column_ag[2]] = template_ag_raw[column_ag[2]].str.split(";")
        # template_ag_raw[column_ag[2]] = template_ag_raw[column_ag[2]] + etl.dict_network.get("store_minigo") + etl.dict_network.get("wh")

        store_minigo = etl.dict_network.get("store_minigo", [])
        wh = etl.dict_network.get("wh", [])

        mask_ok = (
            template_ag_raw["DISCOUNT TYPE"].isin(["1", "2"])
            | (
                template_ag_raw["DISCOUNT TYPE"].eq("3")
                & template_ag_raw["DISCOUNT VALUE"].astype(str).str.contains(
                    r"T|TH", case=False, na=False
                )
            )
        )

        site_lists = (
            template_ag_raw[column_ag[2]]
            .fillna("")
            .astype(str)
            .str.split(";")
            .astype(object)
        )
        site_lists.loc[mask_ok] = site_lists.loc[mask_ok].map(
            lambda sites: sites + store_minigo + wh
        )
        site_lists.loc[~mask_ok] = site_lists.loc[~mask_ok].map(
            lambda sites: sites + store_minigo
        )
        template_ag_raw[column_ag[2]] = site_lists

        template_ag_raw = template_ag_raw.explode(column_ag[2])

        template_ag_raw = template_ag_raw.drop_duplicates()
 
        site_padded = template_ag_raw[column_ag[2]].astype(str).str.zfill(4)
        supplier = template_ag_raw[column_ag[3]].astype(str)
 
        key = supplier + site_padded
        sequence = key.groupby(key).cumcount() + 1
        sequence_str = sequence.apply(lambda x: f"{x:02d}")
 
        column_ag5_value = supplier + site_padded + sequence_str
 
        template_ag_raw = {**template_ag_raw, column_ag[5]: column_ag5_value}
 
        template_ag_raw = pd.DataFrame(template_ag_raw)
 
        template_ag_raw["CONTRACT"] = template_ag_raw.apply(self.get_contract, axis=1)
 
        template_ag_raw = self.fast_stage(template_ag_raw, have_no=True)
 
        self.template_ag_raw = template_ag_raw
 
        return self

    def _create_ag(self) -> "Discount":
        data = self._require_ag_raw()

        template_ag = data[column_ag]

        template_ag = self.fast_stage(template_ag, have_no=True)
 
        self.template_ag = template_ag
 
        return self

    def _update(self, path_report_ag) -> "Discount":
        data = self._require_ag_raw().copy()
        report_paths = path_report_ag if isinstance(path_report_ag, (list, tuple, set)) else [path_report_ag]
        report = pd.concat(
            [pd.read_excel(path, dtype=str) for path in report_paths],
            ignore_index=True,
        )

        group_columns = [
            'ACTION', 'DEPARTMENT', 'SITE', 'SUPPLIER', 'CONTRACT', 'AG NO', 
            'AG DESCRIPTION', 'AG START DATE', 'AG END DATE', 'GOLD CODE', 
            'LV', 'ARTICLE START DATE', 'ARTICLE END DATE'
        ]
        
        group_report_columns = [
            'ACTION', 'DEPT', 'SITE', 'SUPPLIER_CODE', 'COMERCIAL_CONTRACT', 'AGNO1', 
            'AG_DESC', 'AG_START_DATE', 'AG_END_DATE', 'ARTICLE_CODE', 
            'LV', 'ARTICLE_START_DATE', 'ARTICLE_END_DATE'
        ]

        required_report_columns = ["ERRORMESS", "AG_CODE", *group_report_columns]
        missing_columns = [column for column in required_report_columns if column not in report.columns]
        if missing_columns:
            raise ValueError("AG report is missing required columns: " + ", ".join(missing_columns))

        error_mask = report["ERRORMESS"].fillna("").astype(str).str.strip().ne("")
        self.report_err = report.loc[error_mask].copy()
        report = report.loc[~error_mask].copy()

        data["KEY"] = data[group_columns].astype(str).agg("-".join, axis=1)
        report["KEY"] = report[group_report_columns].astype(str).agg("-".join, axis=1)

        report_key = dict(zip(report["KEY"], report["AG_CODE"]))

        data["AG CODE"] = data["KEY"].map(report_key)

        self.template_ag_raw = data.loc[
            data["AG CODE"].fillna("").str.strip().ne("")
        ]

        return self

    def _create_dc(self) -> "Discount":
        ag_raw = self._require_ag_raw()
        self.template_dc_free = None
        self.template_dc_money = None
        
        mask_free = ag_raw["DISCOUNT TYPE"] == "3"

        template_dc_free = {
            column_dc[0]:"0",
            column_dc[1]:ag_raw.loc[mask_free, "SITE"],
            column_dc[2]:ag_raw.loc[mask_free, "SUPPLIER"],
            column_dc[3]:ag_raw.loc[mask_free, "CONTRACT"],
            column_dc[4]:ag_raw.loc[mask_free, "AG CODE"],
            column_dc[5]:ag_raw.loc[mask_free, "AG DESCRIPTION"],
            column_dc[6]:ag_raw.loc[mask_free, "AG CODE"],
            column_dc[7]:"501",
            column_dc[8]:"1",
            column_dc[9]:"2",
            column_dc[10]:"0",
            column_dc[11]:"20",
            column_dc[12]:"20",
            column_dc[13]:ag_raw.loc[mask_free, "ARTICLE START DATE"],
            column_dc[14]:ag_raw.loc[mask_free, "ARTICLE END DATE"],
            column_dc[15]:"3",
            column_dc[16]:"0",
            column_dc[17]:"1",
            column_dc[18]:ag_raw.loc[mask_free, "GOLD CODE"],
            column_dc[19]:ag_raw.loc[mask_free, "LV"],
            column_dc[24]:ag_raw.loc[mask_free, "RAW START DATE"],
            column_dc[25]:ag_raw.loc[mask_free, "RAW END DATE"],
            "VALUE FOR FREE":ag_raw.loc[mask_free, "DISCOUNT VALUE"],
            column_dc[26]:""
        }
        
        template_dc_free = pd.DataFrame(template_dc_free)

        if not template_dc_free.empty:
        
            free_parts = template_dc_free["VALUE FOR FREE"].fillna("").astype(str).str.split("+", n=1, expand=True)
            template_dc_free[column_dc[20]] = free_parts[0]
            template_dc_free[column_dc[22]] = free_parts[1].fillna("")
            
            template_dc_free[column_dc[21]] = template_dc_free[column_dc[20]].map(
                lambda x: "41" if "T" in str(x).upper() else "1"
            )
            
            template_dc_free[column_dc[23]] = template_dc_free[column_dc[22]].map(
                lambda x: "41" if "T" in str(x).upper() else "1"
            )
            
            # wh_network = set(self.etl.dict_network.get("wh", []))

            # template_dc_free[column_dc[21]] = template_dc_free.apply(
            #     lambda row: "41"
            #     if (
            #         (pd.notna(row[column_dc[20]]) and ("TH" in row[column_dc[20]] or "T" in row[column_dc[20]]))
            #         or row[column_dc[1]] in wh_network
            #     )
            #     else "1",
            #     axis=1,
            # )

            # template_dc_free[column_dc[23]] = template_dc_free.apply(
            #     lambda row: "41"
            #     if (
            #         (pd.notna(row[column_dc[22]]) and ("TH" in row[column_dc[22]] or "T" in row[column_dc[22]]))
            #         or row[column_dc[1]] in wh_network
            #     )
            #     else "1",
            #     axis=1,
            # )

            template_dc_free[column_dc[20]] = (
                template_dc_free[column_dc[20]]
                .str.replace(r"th|t", "", case=False, regex=True)
                .str.strip()
            )
            
            template_dc_free[column_dc[22]] = (
                template_dc_free[column_dc[22]]
                .str.replace(r"th|t", "", case=False, regex=True)
                .str.strip()
            )
    
            template_dc_free = template_dc_free[column_dc]
            template_dc_free = self.fast_stage(template_dc_free, have_no=True)
            self.template_dc_free = template_dc_free

        template_dc_money = {
            column_dc[0]:"0",
            column_dc[1]:ag_raw.loc[~mask_free, "SITE"],
            column_dc[2]:ag_raw.loc[~mask_free, "SUPPLIER"],
            column_dc[3]:ag_raw.loc[~mask_free, "CONTRACT"],
            column_dc[4]:ag_raw.loc[~mask_free, "AG CODE"],
            column_dc[5]:ag_raw.loc[~mask_free, "AG DESCRIPTION"],
            column_dc[6]:ag_raw.loc[~mask_free, "AG CODE"],
            column_dc[7]:"501",
            column_dc[8]:"1",
            column_dc[9]:"2",
            column_dc[10]:"0",
            column_dc[11]:"20",
            column_dc[12]:"20",
            column_dc[13]:ag_raw.loc[~mask_free, "ARTICLE START DATE"],
            column_dc[14]:ag_raw.loc[~mask_free, "ARTICLE END DATE"],
            column_dc[15]:ag_raw.loc[~mask_free, "DISCOUNT TYPE"],
            column_dc[16]:"0",
            column_dc[17]:"1",
            column_dc[18]:"",
            column_dc[19]:"",
            column_dc[20]:"",
            column_dc[21]:"",
            column_dc[22]:"",
            column_dc[23]:"",
            column_dc[24]:ag_raw.loc[~mask_free, "ARTICLE START DATE"],
            column_dc[25]:ag_raw.loc[~mask_free, "ARTICLE END DATE"],
            column_dc[26]:""
        }

        template_dc_money = pd.DataFrame(template_dc_money)
        
        if not template_dc_money.empty:
            
            template_dc_money = template_dc_money[column_dc]
    
            template_dc_money = self.fast_stage(template_dc_money, have_no=True)
     
            self.template_dc_money = template_dc_money
 
        return self

    def _create_de(self) -> "Discount":
        raw = self._require_ag_raw()
        
        raw = raw[raw["DISCOUNT TYPE"] != "3"]
        
        template_de = {
            column_de[0]:"0",
            column_de[1]:raw["SITE"],
            column_de[2]:raw["AG CODE"],
            column_de[3]:raw["AG CODE"],
            column_de[4]:raw["GOLD CODE"],
            column_de[5]:raw["LV"],
            column_de[6]:raw["DISCOUNT VALUE"],
            column_de[7]:raw["RAW START DATE"],
            column_de[8]:raw["RAW END DATE"],
            column_de[9]:"",
            column_de[10]:"",
        }
        
        template_de = pd.DataFrame(template_de)
        
        template_de["VALUE ON INVOICE"] = (
            template_de["VALUE ON INVOICE"].fillna("").astype(str).str.replace("%", "").str.strip()
        )

        template_de = template_de[column_de]

        template_de = self.fast_stage(template_de, have_no=True)
 
        self.template_de = template_de
 
        return self

class SalePrice(StageMixin):
    def __init__(self, etl: "Template_ETL"):
        self.src = etl.src_listoff
        self.src_attr = etl.src_attr
        self.etl = etl
 
        self.template_sp: Optional[pd.DataFrame] = None
        self.template_attr: Optional[pd.DataFrame] = None

    def _create_sp(self) -> "SalePrice":
        data = self.src.copy()
        template_sp = {
            column_sale[0] : "0",
            column_sale[1] : data["GOLD CODE"],
            column_sale[2] : data["SV"],
            column_sale[3] : data["PROMOTION SALE PRICE"],
            column_sale[4] : data["PRICELIST"],
            "PRICELIST CODE" : data["PRICELIST CODE"],
            column_sale[5] : data["SP START DATE"].dt.strftime("%d/%m/%Y"),
            column_sale[6] : data["SP END DATE"].dt.strftime("%d/%m/%Y"),
            column_sale[7] : data["SALE VAT"].map(VAT),
            column_sale[8] : ""
        }

        template_sp = pd.DataFrame(template_sp)
        template_sp["PRICELIST"] = template_sp["PRICELIST"].str.split(";")
        template_sp = template_sp.explode("PRICELIST")
        def _get_plcode(x):
            PRICELIST = x["PRICELIST"]
            CODE = x["PRICELIST CODE"]
        
            if len(PRICELIST) == 3:
                return f"{CODE}0{PRICELIST}"
            if len(PRICELIST) == 4:
                return f"{CODE}{PRICELIST}"
        
        template_sp["PRICELIST"] = template_sp.apply(_get_plcode, axis=1)
        template_sp = template_sp[column_sale]

        template_sp = self.fast_stage(template_sp, have_no=True)

        self.template_sp = template_sp

        return self

    def _create_attr(self) -> "SalePrice":
        if self.src_attr is None:
            raise ValueError("Attribute data has not been loaded.")
        data = self.src_attr.copy()
        note_count = data.groupby(["GOLD CODE", "SV"], dropna=False)[
            "GOLD CODE"
        ].transform("size")

        template_attr = pd.DataFrame({
            column_attr[0]: "0",
            column_attr[1]: data["GOLD CODE"],
            column_attr[2]: data["SV"],
            column_attr[3]: data["CLASS"],
            column_attr[4]: f"{self.etl.cata}D",
            column_attr[5]: data["Alphanum"],
            column_attr[6]: "",
            column_attr[7]: "",
            column_attr[8]: "",
            column_attr[9]: data["START DATE"],
            column_attr[10]: data["END DATE"],
            column_attr[11]: note_count,
        })

        self.template_attr = self.fast_stage(template_attr[column_attr], have_no=True)
        return self
