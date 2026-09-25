from collections import Counter
from decimal import Decimal, InvalidOperation
import os
from pathlib import Path
import re
import secrets
import tempfile
from typing import List, Optional

import pandas as pd
from openpyxl import load_workbook
from src.workbook_state import append_sitegroup_history, workbook_write_lock

from src.constant.required import (
    date_columns_plan,
    department,
    required_cm,
    required_network,
    required_src,
    required_stage1,
    required_stage2,
    required_wh_discount,
    VAT,
)


class Template_ETL:
    GROUP_COLS: List[str] = ["FILE NAME", "GOLD CODE", "LV", "LU"]
    DISCOUNT_NUMBER = r"\d+(?:[.,]\d+)?"
    VALID_DISCOUNT = re.compile(
        rf"^(?:{DISCOUNT_NUMBER}|{DISCOUNT_NUMBER}%|"
        rf"{DISCOUNT_NUMBER}(?:T|TH)\+{DISCOUNT_NUMBER}(?:T|TH)|"
        rf"{DISCOUNT_NUMBER}\+{DISCOUNT_NUMBER})$",
        re.IGNORECASE,
    )
    NON_WAREHOUSE_DISCOUNT = re.compile(
        rf"^{DISCOUNT_NUMBER}\+{DISCOUNT_NUMBER}$"
    )
    ATTRIBUTE_COLUMNS = (
        "GOLD CODE",
        "SV",
        "START DATE",
        "END DATE",
        "REGION (NORTH/SOUTH/CENTER/ALL)",
        "PAGE",
        "THEMATIC",
        "POSITION",
    )
    ATTRIBUTE_CLASS_RULES = (
        (re.compile(r"(?i)\b(front\s*page|back\s*page|unbeat|hero)\b"), "HEROP"),
        (re.compile(r"(?i)\b(cata|catalog(?:ue)?|fair|member\s*price|banner|exclusive\s*pack|family|other|normal|the\s*1)\b"), "CATAP"),
        (re.compile(r"(?i)\b(comple(?:mentary)?|comple)\b"), "COMPLEP"),
        (re.compile(r"(?i)\bbuy\s*more\s*save\s*more\b"), "STARP"),
    )
    CATEGORY_RULES = (
        (re.compile(r"(?i)\b(front\s*page|back\s*page|unbeat|hero)\b"), "HERO"),
        (re.compile(r"(?i)\b(?:star|buy\s*more\s*save\s*more)\b"), "STAR"),
        (re.compile(r"(?i)\bmodel\b"), "MODEL"),
        (
            re.compile(
                r"(?i)\b(cata|catalog(?:ue)?|fair|member\s*price|banner|exclusive\s*pack|family|other|normal|the\s*1)\b"
            ),
            "CATA",
        ),
        (re.compile(r"(?i)\b(comple(?:mentary)?|comple)\b"), "COMPLE"),
    )
    ATTRIBUTE_MARKETING_ERROR = "Vui lòng bổ sung prefix cho ATTRIBUTE đặc biệt"
    ATTRIBUTE_CLASS_ERROR = "POSITION không thể chuyển đổi thành CLASS"
    NORMAL_PURCHASE_PRICE_ERROR = "NORMAL PURCHASE PRICE không thể chuyển đổi thành số"
    DISCOUNT_VALUE_ERROR = "DISCOUNT (% OR VALUE) không thể chuyển đổi thành số"
    DISCOUNT_PERCENTAGE_LIMIT_ERROR = "DISCOUNT (% OR VALUE) không được vượt quá 100%."
    ALLOWED_DIFFERENT_DISCOUNT_TYPE_PAIRS = {
        frozenset(("1", "3")),
        frozenset(("2", "3")),
    }

    def __init__(
        self,
        path_src,
        path_sitegroup=None,
        path_plan=None,
        path_attribute=None,
        non_suggested_sitegroup_codes=None,
        excluded_sitegroup_codes=None,
        exception_discount_gold_codes=None,
        check_attribute=False,
    ):
        if path_src is None:
            self.path_src: tuple[Path, ...] = ()
        elif isinstance(path_src, (str, Path)):
            self.path_src = (Path(path_src),)
        else:
            self.path_src = tuple(Path(path) for path in path_src)
        if not self.path_src:
            self.path_src = ()
        self.path_attribute = Path(path_attribute) if path_attribute is not None else None
        self.path_sitegroup = Path(path_sitegroup) if path_sitegroup is not None else None
        self.path_plan = Path(path_plan) if path_plan is not None else None
        self.check_attribute = bool(check_attribute)

        self.dict_network: dict = dict()

        self.sitegroup: dict = dict()
        self.sitegroup_members: dict[str, tuple[str, ...]] = dict()
        self.master_sitegroup_codes: set[str] = set()
        self.activated_sitegroup_codes: set[str] = set()
        # These codes remain valid for exact Site Group matching. They are
        # reserved only from automatic suggestions.
        reserved_sitegroup_codes = (
            non_suggested_sitegroup_codes
            if non_suggested_sitegroup_codes is not None
            else excluded_sitegroup_codes
        )
        self.non_suggested_sitegroup_codes = {
            str(code).strip()
            for code in (reserved_sitegroup_codes or [])
            if str(code).strip()
        }
        self.exception_discount_gold_codes = {
            str(code).strip()
            for code in (exception_discount_gold_codes or [])
            if str(code).strip()
        }

        self.src: Optional[pd.DataFrame] = None
        self.gold_code_delete: Optional[pd.DataFrame] = None
        self.non_warehouse_src: Optional[pd.DataFrame] = None
        self.src_listoff: Optional[pd.DataFrame] = None
        self.src_attr: Optional[pd.DataFrame] = None
        self.attribute_sheet_name: str | None = None
        self.attribute_header_row: int | None = None

        self.plan: Optional[pd.DataFrame] = None

        self.cata: str = str()
        # The value is keyed by source filename because a run may contain one
        # source workbook per department.
        self.dept: dict[str, str] = {}
        self.cata_description: str = str()
        self.cata_period: str = str()
        self.should_generate_so_sitegroup = True
        self.should_generate_so = True

    @staticmethod
    def _department_code(value) -> str:
        """Convert the source DEPARTMENT value to the code used by templates."""
        department = "" if pd.isna(value) else str(value).strip()
        return department[:4].lstrip("0") or "0"

    @staticmethod
    def _percentage_discount_exceeds_limit(value: str) -> bool:
        """Return whether any percentage component is greater than 100."""
        for component in str(value).split("+"):
            if not component.endswith("%"):
                continue
            try:
                percentage = float(component[:-1].replace(",", "."))
            except ValueError:
                continue
            if percentage > 100:
                return True
        return False

    def _load_source_metadata(
        self,
        sheet_names: dict[Path, str] | None = None,
    ) -> dict[Path, pd.DataFrame]:
        """Read metadata from every source and ensure they form one catalogue run."""
        if not self.path_src:
            raise ValueError("At least one Gold Promo source file is required.")

        comparison_columns = [
            "CATALOGUE START DAY", "CATALOGUE START MONTH", "CATALOGUE START YEAR",
            "CATALOGUE END DAY", "CATALOGUE END MONTH", "CATALOGUE END YEAR",
            "CATALOGUE",
        ]
        required_columns = [*comparison_columns, "DEPARTMENT", "CATALOGUE DESCRIPTION"]
        metadata = {
            path: pd.read_excel(
                path,
                nrows=2,
                dtype=str,
                sheet_name=(sheet_names or {}).get(path, "Template"),
            )
            for path in self.path_src
        }
        for path, data in metadata.items():
            self._check_required_columns(data, required_columns)
            self._check_required_data(data, required_columns)
            if data.empty:
                raise ValueError(f"Source file has no metadata rows: {path.name}")
            catalogue_start, catalogue_end = self._combine_date_columns(data, "CATALOGUE")
            if catalogue_start.isna().any() or catalogue_end.isna().any():
                raise ValueError(f"Invalid catalogue start/end date in source metadata: {path.name}")

        first_path, first_data = next(iter(metadata.items()))
        reference = first_data.loc[:, comparison_columns].fillna("").astype(str).reset_index(drop=True)
        mismatched = [
            path.name
            for path, data in metadata.items()
            if not data.loc[:, comparison_columns].fillna("").astype(str).reset_index(drop=True).equals(reference)
        ]
        if mismatched:
            raise ValueError(
                "The following metadata columns must be identical in every source file: "
                + ", ".join(comparison_columns)
                + ". Mismatched files: "
                + ", ".join(mismatched)
            )

        self.cata = str(first_data["CATALOGUE"].iat[0]).strip()
        self.cata_description = str(first_data["CATALOGUE DESCRIPTION"].iat[0]).strip()
        self.cata_period = (
            f"{first_data['CATALOGUE START DAY'].iat[0]}/{first_data['CATALOGUE START MONTH'].iat[0]} - "
            f"{first_data['CATALOGUE END DAY'].iat[0]}/{first_data['CATALOGUE END MONTH'].iat[0]}/{first_data['CATALOGUE END YEAR'].iat[0]}"
        )
        self.dept = {
            path.name: self._department_code(data["DEPARTMENT"].str[:4].str.strip("0").iat[0])
            for path, data in metadata.items()
        }
        return metadata

    @staticmethod
    def _check_required_columns(
        data: Optional[pd.DataFrame],
        required: List[str]
    ) -> Optional[pd.DataFrame]:
        dup_cols = data.columns[data.columns.duplicated()].unique().tolist()
        if dup_cols:
            raise ValueError("Dữ liệu có cột trùng lặp.")

        missing_cols = [c for c in required if c not in data.columns]
        if missing_cols:
            raise ValueError(f"Dữ liệu thiếu các cột {', '.join(missing_cols)}")

        return data

    @classmethod
    def _check_required_data(
        cls,
        data: Optional[pd.DataFrame],
        required: List[str]
    ) -> Optional[pd.DataFrame]:
        if data is None or data.empty:
            return data

        empty_values = data.loc[:, required].apply(
            lambda column: column.isna()
            | column.fillna("").astype(str).str.strip().eq("")
        )
        empty_cols = empty_values.columns[empty_values.any()].tolist()

        if empty_cols:
            if "NOTE ERR FROM MASTER DATA" not in data.columns:
                raise ValueError(
                    f"Các cột bắt buộc đang để trống: {', '.join(empty_cols)}"
                )

            for index, row in empty_values.loc[empty_values.any(axis=1)].iterrows():
                missing = row.index[row].tolist()
                cls._append_note_err(
                    data,
                    pd.Index([index]),
                    f"Các cột bắt buộc đang để trống: {', '.join(missing)}",
                )

        return data

    @staticmethod
    def _ensure_note_err(data: pd.DataFrame) -> pd.DataFrame:
        if "NOTE ERR FROM MASTER DATA" not in data.columns:
            data["NOTE ERR FROM MASTER DATA"] = ""
        data["NOTE ERR FROM MASTER DATA"] = data["NOTE ERR FROM MASTER DATA"].fillna("")
        return data

    @staticmethod
    def _append_note_err(data: pd.DataFrame, idx: None, message: str) -> None:
        if not message:
            return

        if idx is None:
            idx = data.index

        data.loc[idx, "NOTE ERR FROM MASTER DATA"] = (
            data.loc[idx, "NOTE ERR FROM MASTER DATA"]
            + data.loc[idx, "NOTE ERR FROM MASTER DATA"].ne("").map({True: " | ", False: ""})
            + message
        )

    def _combine_date_columns(
        self,
        data: pd.DataFrame,
        date_prefix: str,
        error_prefix: str | None = None,
    ) -> tuple[pd.Series, pd.Series]:
        """Combine day/month/year columns and report invalid dates."""
        error_prefix = error_prefix or date_prefix

        def combine(kind: str) -> pd.Series:
            columns = [
                f"{date_prefix} {kind} DAY",
                f"{date_prefix} {kind} MONTH",
                f"{date_prefix} {kind} YEAR",
            ]
            components = data[columns].apply(pd.to_numeric, errors="coerce")
            integer_components = components.notna() & components.eq(components.round())
            valid_components = integer_components.all(axis=1)
            text = components.round().astype("Int64").astype("string")
            result = pd.to_datetime(
                text[columns[0]] + "/" + text[columns[1]] + "/" + text[columns[2]],
                format="%d/%m/%Y",
                errors="coerce",
            )
            invalid = ~valid_components | result.isna()
            if "NOTE ERR FROM MASTER DATA" in data.columns:
                self._append_note_err(
                    data,
                    data.index[invalid],
                    f"{error_prefix} {kind} DATE không hợp lệ.",
                )
            return result

        return combine("START"), combine("END")

    @staticmethod
    def _normalize_decimal_number(value) -> str:
        """Normalize decimal separators used by different regional formats."""
        if pd.isna(value):
            return ""

        text = str(value).strip().replace("\u00a0", "").replace(" ", "")
        comma_position = text.rfind(",")
        dot_position = text.rfind(".")

        if comma_position >= 0 and dot_position >= 0:
            if comma_position > dot_position:
                return text.replace(".", "").replace(",", ".")
            return text.replace(",", "")

        if comma_position >= 0:
            return text.replace(",", ".")

        return text

    def _unique_sorted_sites(self, value) -> tuple:
        return tuple(sorted(set(self._parse_sites(value)), key=self._sort_key))

    @staticmethod
    def _default_blank_discounts(data: pd.DataFrame) -> pd.DataFrame:
        """Treat missing source discounts as an explicit zero percent."""
        column = "DISCOUNT (% OR VALUE)"
        blank = data[column].fillna("").astype(str).str.strip().eq("")
        data.loc[blank, column] = "0%"
        return data

    def _load_network(self) -> "Template_ETL":
        data = pd.read_excel(
            self.path_plan,
            dtype=str,
            sheet_name="network-configure"
        )

        self._check_required_columns(data, required_network)
        self._check_required_data(data, required_network)

        discount_rows = data.loc[data["DISCOUNT"].eq("1")]
        discount_groups = pd.concat([
            discount_rows[[column, "SITE"]].rename(columns={column: "GROUP"})
            for column in ("NATIONAL_SITE", "GROUP_SITE", "REGION_SITE")
        ]).drop_duplicates().dropna()
        discount_network = discount_groups.groupby("GROUP")["SITE"].agg(";".join).to_dict()

        data = data[data['ACTIVE'] == '1']

        NATIONAL_SITE = {
            'GROUP' : data['NATIONAL_SITE'],
            'NETWORK' : data['SITE']
        }

        GROUP_SITE = {
            'GROUP' : data['GROUP_SITE'],
            'NETWORK' : data['SITE']
        }

        REGION_SITE = {
            'GROUP' : data['REGION_SITE'],
            'NETWORK' : data['SITE']
        }

        data = pd.concat([pd.DataFrame(NATIONAL_SITE), pd.DataFrame(GROUP_SITE), pd.DataFrame(REGION_SITE)])

        data = data.drop_duplicates()
        data = data.dropna()

        data = (
            data.groupby("GROUP", as_index=False)["NETWORK"]
            .agg(";".join)
        )

        network = dict(zip(data["GROUP"], data["NETWORK"]))

        store_hyper = str(network.get("8200", "")).split(";")
        store_minigo = str(network.get("8710", "")).split(";")
        store = store_hyper + store_minigo

        wh = str(network.get("8300", "")).split(";")
        wh8 = [s for s in wh if s.startswith("8")]
        wh9 = [s for s in wh if s.startswith("9")]

        self.dict_network = {
            "network": network,
            "DISCOUNT_NETWORK": discount_network,
            "store_hyper": store_hyper,
            "store_minigo": store_minigo,
            "store": store,
            "wh": wh,
            "wh8": wh8,
            "wh9": wh9,
        }

        # Stage 2 intentionally loads the source before the network.  Run its
        # network-dependent validation once the network dictionary is ready.
        if self.src_listoff is not None:
            self._validate_sale_pricelists(self.src_listoff)

        return self

    def _validate_sale_pricelists(self, data: pd.DataFrame) -> pd.DataFrame:
        """Append Stage 2 errors for stores absent from the configured network."""
        self._ensure_note_err(data)
        valid_network = set(self.dict_network.get("store", []))

        for index, value in data["PRICELIST"].items():
            pricelist = self._parse_sites(value)
            invalid = [site for site in pricelist if site not in valid_network]
            if invalid:
                self._append_note_err(
                    data,
                    pd.Index([index]),
                    f"Cửa hàng {';'.join(invalid)} không tồn tại.",
                )

        return data

    @staticmethod
    def _normalize_vat(value) -> str:
        """Convert Excel percentage fractions to the supported VAT labels."""
        if pd.isna(value):
            return ""
        text = str(value).strip().upper()
        if text in VAT or not text:
            return text
        try:
            fraction = Decimal(text.replace(",", "."))
        except InvalidOperation:
            return text
        if not fraction.is_finite():
            return text
        return {
            Decimal("0"): "0%",
            Decimal("0.05"): "5%",
            Decimal("0.08"): "8%",
            Decimal("0.1"): "10%",
        }.get(fraction, text)

    def _validate_stage2_sale_values(self, data: pd.DataFrame) -> pd.DataFrame:
        """Validate VAT and normalize valid Sale Price values to integers."""
        self._ensure_note_err(data)

        data["SALE VAT"] = data["SALE VAT"].map(self._normalize_vat)
        invalid_vat = ~data["SALE VAT"].isin(VAT)
        self._append_note_err(
            data,
            data.index[invalid_vat],
            "SALE VAT chỉ được phép là 0%, 5%, 8%, 10%, KKKT hoặc KCT.",
        )

        price_text = data["PROMOTION SALE PRICE"].fillna("").astype(str).str.strip()
        normalized_price = price_text.map(self._normalize_decimal_number)
        numeric_price = pd.to_numeric(normalized_price, errors="coerce")
        populated_price = price_text.ne("")
        invalid_price = populated_price & (
            numeric_price.isna()
            | (numeric_price - numeric_price.round()).abs().gt(1e-9)
        )
        self._append_note_err(
            data,
            data.index[invalid_price],
            "PROMOTION SALE PRICE phải là số nguyên.",
        )

        valid_price = populated_price & ~invalid_price
        data.loc[valid_price, "PROMOTION SALE PRICE"] = (
            numeric_price.loc[valid_price].round().astype("int64").astype(str)
        )
        return data

    def _load_attribute(self, sheet_name: str | None = None) -> "Template_ETL":
        raw_data = pd.read_excel(
            self.path_attribute,
            dtype=str,
            header=None,
            sheet_name=sheet_name,
        )
        required_columns = set(self.ATTRIBUTE_COLUMNS)
        header_row = next(
            (
                index
                for index, row in raw_data.head(10).iterrows()
                if required_columns.issubset(
                    {str(value).strip() for value in row if pd.notna(value) and str(value).strip()}
                )
            ),
            None,
        )
        if header_row is None:
            raise ValueError(
                "Attribute file is missing required columns: "
                + ", ".join(self.ATTRIBUTE_COLUMNS)
            )

        data = pd.read_excel(
            self.path_attribute,
            dtype=str,
            header=header_row,
            sheet_name=sheet_name,
        )
        # data.columns = [str(column).strip() for column in data.columns]
        # self._check_required_columns(data, list(self.ATTRIBUTE_COLUMNS))
        data = data.loc[:, self.ATTRIBUTE_COLUMNS].dropna(how="all").copy()
        data = self._ensure_note_err(data)
        required_attribute_data = [
            column
            for column in self.ATTRIBUTE_COLUMNS
            if column not in {"START DATE", "END DATE"}
        ]
        self._check_required_data(data, required_attribute_data)
        data["_SOURCE_ROW"] = data.index + header_row + 2

        def attribute_class(value) -> str | None:
            text = "" if pd.isna(value) else str(value).strip()
            for pattern, attribute in self.ATTRIBUTE_CLASS_RULES:
                if pattern.search(text):
                    return attribute
            return None

        data["CLASS"] = data["POSITION"].map(attribute_class)
        invalid_class = data["CLASS"].isna()
        self._append_note_err(
            data,
            data.index[invalid_class],
            self.ATTRIBUTE_CLASS_ERROR,
        )
        data["Alphanum"] = (
            data["POSITION"].astype(str).str.strip()
            + ".P."
            + data["PAGE"].astype(str).str.strip()
            + "."
            + data["REGION (NORTH/SOUTH/CENTER/ALL)"].astype(str).str.strip()
        )

        if self.plan is None or self.plan.empty:
            raise ValueError(f"Không tìm thấy CATALOGUE {self.cata} trong Master data.")

        catalogue_start_date = pd.to_datetime(
            self.plan["CATALOGUE START DATE"].iat[0],
            errors="coerce",
            dayfirst=True,
        )
        catalogue_end_date = pd.to_datetime(
            self.plan["CATALOGUE END DATE"].iat[0],
            errors="coerce",
            dayfirst=True,
        )
        if pd.isna(catalogue_start_date) or pd.isna(catalogue_end_date):
            raise ValueError(
                f"CATALOGUE {self.cata} có START DATE hoặc END DATE không hợp lệ trong Master data."
            )

        data["START DATE"] = catalogue_start_date.strftime("%d/%m/%Y")
        data["END DATE"] = catalogue_end_date.strftime("%d/%m/%Y")

        self.src_attr = data.drop_duplicates().reset_index(drop=True)
        self.attribute_sheet_name = sheet_name
        self.attribute_header_row = header_row + 1

        return self
        
    def _load_src(self, validate_source: bool = True) -> "Template_ETL":
        self._load_source_metadata()
        sources = []
        deleted_sources = []
        for path in self.path_src:
            data = pd.read_excel(path, header=6, dtype=str, sheet_name="Template")
            data = data.drop(columns=["NOTE ERR FROM MASTER DATA"], errors="ignore")
            for col in ["NOTE ERR FROM MASTER DATA", "SITE GROUP", "SO", "STRUCTURE"]:
                if col not in data.columns:
                    data[col] = ""
            self._check_required_columns(data, required_cm)
            if validate_source:
                delete_mask = data["ATTRIBUTE MARKETING"].fillna("").astype(str).str.contains(
                    "delete", case=False, na=False
                )
                deleted = data.loc[delete_mask, ["GOLD CODE", "LV", "SO"]].copy()
                deleted["FILE NAME"] = path.name
                deleted["STRUCTURE"] = self.dept[path.name]
                deleted_sources.append(deleted)
                data = data.loc[~delete_mask].copy()
            data = self._default_blank_discounts(data)
            required_source_data = [
                column for column in required_stage1 if column != "FREE PRODUCT"
            ]
            if not self.check_attribute:
                required_source_data = [
                    column for column in required_stage1
                    if column not in {"FREE PRODUCT", "ATTRIBUTE MARKETING"}
                ]
            if validate_source:
                self._check_required_data(data, required_source_data)
            # Temporarily disabled to avoid reopening and scanning every source
            # workbook with openpyxl. Keep the helper available for later use.
            # self._restore_percentage_cells(data, path)
            if not validate_source:
                data.columns = [str(col).replace("-RECOMMENDATION QUALITY", "") for col in data.columns]
                data["FILE NAME"] = path.name
                data["_SOURCE_ROW"] = data.index + 8
                sources.append(data)
                continue

            if self.check_attribute:
                converted_attribute = pd.Series(pd.NA, index=data.index, dtype="string")
                attribute_text = data["ATTRIBUTE MARKETING"].fillna("").astype(str).str.strip()
                for pattern, category in self.CATEGORY_RULES:
                    matched = converted_attribute.isna() & attribute_text.map(
                        lambda text: bool(pattern.search(text))
                    )
                    converted_attribute.loc[matched] = category

                invalid_attribute = converted_attribute.isna()
                self._append_note_err(
                    data,
                    data.index[invalid_attribute],
                    self.ATTRIBUTE_MARKETING_ERROR,
                )
                data.loc[~invalid_attribute, "ATTRIBUTE MARKETING"] = converted_attribute.loc[
                    ~invalid_attribute
                ]

            data["PURCHASE VAT"] = data["PURCHASE VAT"].map(self._normalize_vat)
            self._append_note_err(
                data,
                data.index[~data["PURCHASE VAT"].isin(VAT)],
                "PURCHASE VAT chỉ được phép là 0%, 5%, 8%, 10%, KKKT hoặc KCT.",
            )

            normalized_purchase_price = data["NORMAL PURCHASE PRICE"].map(
                self._normalize_decimal_number
            )
            normal_purchase_price = pd.to_numeric(
                normalized_purchase_price, errors="coerce"
            )
            invalid_normal_purchase_price = normal_purchase_price.isna()
            self._append_note_err(
                data,
                data.index[invalid_normal_purchase_price],
                self.NORMAL_PURCHASE_PRICE_ERROR,
            )
            data["NORMAL PURCHASE PRICE"] = normal_purchase_price.astype(float)

            discount_text = (
                data["DISCOUNT (% OR VALUE)"]
                .fillna("")
                .astype(str)
                .str.replace(r"\s+", "", regex=True)
            )
            plain_discount = ~discount_text.str.contains(r"[+%]", regex=True, na=False)
            numeric_discount = pd.to_numeric(
                discount_text.where(plain_discount).map(self._normalize_decimal_number),
                errors="coerce",
            )
            invalid_discount = ~discount_text.str.fullmatch(
                self.VALID_DISCOUNT, na=False
            )
            self._append_note_err(
                data,
                data.index[invalid_discount],
                self.DISCOUNT_VALUE_ERROR,
            )
            percentage_over_limit = discount_text.map(
                self._percentage_discount_exceeds_limit
            )
            self._append_note_err(
                data,
                data.index[percentage_over_limit],
                self.DISCOUNT_PERCENTAGE_LIMIT_ERROR,
            )
            valid_plain_discount = plain_discount & ~invalid_discount
            data["DISCOUNT (% OR VALUE)"] = data[
                "DISCOUNT (% OR VALUE)"
            ].astype(object)
            data.loc[~plain_discount, "DISCOUNT (% OR VALUE)"] = discount_text.loc[
                ~plain_discount
            ]
            data.loc[
                valid_plain_discount, "DISCOUNT (% OR VALUE)"
            ] = numeric_discount.loc[valid_plain_discount].astype(float)
            data.columns = [str(col).replace("-RECOMMENDATION QUALITY", "") for col in data.columns]
            data["FILE NAME"] = path.name
            data["_SOURCE_ROW"] = data.index + 8
            sources.append(data)

        self.src = pd.concat(sources, ignore_index=True)
        self.gold_code_delete = (
            pd.concat(deleted_sources, ignore_index=True).drop_duplicates().reset_index(drop=True)
            if deleted_sources else None
        )
        self.should_generate_so_sitegroup = True
        self.should_generate_so = True

        return self

    @staticmethod
    def _restore_percentage_cells(data: pd.DataFrame, path: Path) -> None:
        """Restore selected percentage-formatted source cells to ``50%`` text.

        Pandas returns an Excel percentage cell such as ``50%`` as ``0.5``
        even with ``dtype=str``.  Read the source cell format so validation
        and template generation receive the user-facing percentage value.
        """
        percentage_columns = (
            "DISCOUNT (% OR VALUE)",
            "% DELIVERY 1",
            "% DELIVERY 2",
            "% DELIVERY 3",
        )
        target_columns = [column for column in percentage_columns if column in data.columns]
        if not target_columns:
            return
        workbook = load_workbook(path, read_only=True, data_only=True)
        try:
            worksheet = workbook["Template"]
            headers = [cell.value for cell in next(worksheet.iter_rows(min_row=7, max_row=7))]
            for column in target_columns:
                if column not in headers:
                    continue
                column_index = headers.index(column) + 1
                for dataframe_index, excel_row in zip(data.index, range(8, len(data) + 8)):
                    cell = worksheet.cell(excel_row, column_index)
                    if "%" not in str(cell.number_format) or not isinstance(cell.value, (int, float)):
                        continue
                    data.at[dataframe_index, column] = f"{cell.value * 100:g}%"
        finally:
            workbook.close()

    def _load_src_wh_discount(self) -> "Template_ETL":
        """Load only the fields needed by the standalone WH Discount flow."""
        self._load_source_metadata()
        sources = []
        for path in self.path_src:
            data = pd.read_excel(path, header=6, dtype=str, sheet_name="Template")
            data = data.drop(columns=["NOTE ERR FROM MASTER DATA"], errors="ignore")
            data["NOTE ERR FROM MASTER DATA"] = ""
            self._check_required_columns(data, required_wh_discount)
            data = self._default_blank_discounts(data)
            self._check_required_data(data, required_wh_discount)
            # Temporarily disabled to avoid reopening and scanning every source
            # workbook with openpyxl. Keep the helper available for later use.
            # self._restore_percentage_cells(data, path)

            discount_text = (
                data["DISCOUNT (% OR VALUE)"]
                .fillna("")
                .astype(str)
                .str.replace(r"\s+", "", regex=True)
            )
            plain_discount = ~discount_text.str.contains(r"[+%]", regex=True, na=False)
            numeric_discount = pd.to_numeric(
                discount_text.where(plain_discount).map(self._normalize_decimal_number),
                errors="coerce",
            )
            invalid_discount = ~discount_text.str.fullmatch(self.VALID_DISCOUNT, na=False)
            self._append_note_err(
                data, data.index[invalid_discount], self.DISCOUNT_VALUE_ERROR
            )
            percentage_over_limit = discount_text.map(self._percentage_discount_exceeds_limit)
            self._append_note_err(
                data,
                data.index[percentage_over_limit],
                self.DISCOUNT_PERCENTAGE_LIMIT_ERROR,
            )
            valid_plain_discount = plain_discount & ~invalid_discount
            data["DISCOUNT (% OR VALUE)"] = data["DISCOUNT (% OR VALUE)"].astype(object)
            data.loc[~plain_discount, "DISCOUNT (% OR VALUE)"] = discount_text.loc[
                ~plain_discount
            ]
            data.loc[valid_plain_discount, "DISCOUNT (% OR VALUE)"] = (
                numeric_discount.loc[valid_plain_discount].astype(float)
            )

            # Keep the original purchase field; Discount uses its own expansion.
            data["PURCHASE NETWORK EXPANDED"] = data["PURCHASE NETWORK"]
            data["DISCOUNT_NETWORK_EXPANDED"] = data["PURCHASE NETWORK"].map(
                self._normalize_network_punctuation
            ).map(self._extract_discount_network)
            discount_wh = set(self._parse_sites(
                self.dict_network.get("DISCOUNT_NETWORK", {}).get("8300", "")
            ))
            for index, expanded in data["DISCOUNT_NETWORK_EXPANDED"].items():
                sites = set(self._parse_sites(expanded))
                invalid_sites = sorted(sites - discount_wh, key=self._sort_key)
                if not sites:
                    self._append_note_err(
                        data, pd.Index([index]),
                        "PURCHASE NETWORK sau khi expand theo DISCOUNT_NETWORK bị rỗng.",
                    )
                elif invalid_sites:
                    self._append_note_err(
                        data, pd.Index([index]),
                        "Site " + ";".join(invalid_sites)
                        + " không thuộc WH (8300) của DISCOUNT_NETWORK.",
                    )
            data["STRUCTURE"] = self.dept[path.name]
            data["FILE NAME"] = path.name
            data["_SOURCE_ROW"] = data.index + 8
            data = self._convert_date(data)
            sources.append(data)

        self.src = pd.concat(sources, ignore_index=True)
        self.should_generate_so_sitegroup = False
        self.should_generate_so = False
        return self

    def clear_so_and_sitegroup(self) -> "Template_ETL":
        """Clear prior SITE GROUP and SO values for a fresh Stage 1 run."""
        if self.src is not None:
            self.src[["SITE GROUP", "SO"]] = ""
        self.should_generate_so_sitegroup = True
        self.should_generate_so = True
        return self

    def clear_sitegroup_keep_so(self) -> "Template_ETL":
        """Keep complete source SO values while forcing Site Group resolution."""
        if self.src is not None:
            self.src["SITE GROUP"] = ""
        self.should_generate_so = False
        self.should_generate_so_sitegroup = True
        return self

    def has_complete_so(self) -> bool:
        """Return whether every loaded source row contains a non-blank SO."""
        if self.src is None or self.src.empty or "SO" not in self.src.columns:
            return False
        return bool(self.src["SO"].fillna("").astype(str).str.strip().ne("").all())

    def prepare_check_oa_without_validation(self) -> "Template_ETL":
        """Build only the derived fields required by Check OA and Site Group."""
        if self.src is None:
            raise ValueError("Stage 1 source data has not been loaded.")
        data = self.src
        for column in ("PURCHASE NETWORK", "GOLD PROMO NETWORK"):
            data[column] = data[column].map(self._normalize_network_punctuation)
        data["PURCHASE NETWORK EXPANDED"] = data["PURCHASE NETWORK"].map(self._extract_network)
        data["DISCOUNT_NETWORK_EXPANDED"] = data["PURCHASE NETWORK"].map(self._extract_discount_network)
        data["GOLD PROMO NETWORK EXPANDED"] = data["GOLD PROMO NETWORK"].map(self._extract_network)
        for column in ("PURCHASE NETWORK EXPANDED", "GOLD PROMO NETWORK EXPANDED"):
            data[column] = data[column].map(self._sort_network)
        data["STRUCTURE"] = data["FILE NAME"].map(self.dept)
        self._ensure_note_err(data)
        data["PP START DATE"], data["PP END DATE"] = self._combine_date_columns(data, "PP")
        data["COMMERCIAL CONTRACT"] = (
            data["COMMERCIAL CONTRACT"].fillna("").astype(str).map(self._contract_checking)
        )
        data["SITE GROUP"] = ""
        self.src = data
        self.should_generate_so = False
        self.should_generate_so_sitegroup = True
        return self

    def _load_src_listoff(
        self,
        sheet_names: dict[Path, str] | None = None,
    ) -> "Template_ETL":
        metadata = self._load_source_metadata(sheet_names)
        sources = []
        period_columns = ['SP START DAY', 'SP START MONTH', 'SP START YEAR', 'SP END DAY', 'SP END MONTH', 'SP END YEAR']
        required = [
            "GOLD CODE",
            "SV",
            "PRICELIST",
            "SALE VAT",
        ]
        for path in self.path_src:
            data = pd.read_excel(
                path,
                header=6,
                dtype=str,
                sheet_name=(sheet_names or {}).get(path, "Template"),
            )
            sale_period = metadata[path]
            data = data.drop(columns=["NOTE ERR FROM MASTER DATA"], errors="ignore")
            data["NOTE ERR FROM MASTER DATA"] = ""
            self._check_required_columns(data, required_stage2)
            self._check_required_columns(sale_period, period_columns)
            self._check_required_data(data, required)
            self._check_required_data(sale_period, period_columns)
            self._validate_stage2_sale_values(data)

            data[period_columns] = sale_period.loc[sale_period.index[0], period_columns].values
    
            pricelist = data["PRICELIST"].fillna("").str.strip().str[:4]
    
            valid_pricelist = ["1090", "2010", "2030", "2050"]
    
            self._append_note_err(
            data,
            data.index[
                data["PRICELIST"].fillna("").str.strip().ne("")
                & ~pricelist.isin(valid_pricelist)
            ],
            "PRICELIST phải bắt đầu bằng 1090, 2010, 2030 hoặc 2050."
            )
    
            mask_normal = pricelist.eq("1090")
    
            self._append_note_err(
            data,
            data.index[
                mask_normal
                & data["NORMAL SALE PRICE"].fillna("").str.strip().eq("")
            ],
            "NORMAL SALE PRICE không được để trống khi PRICELIST bắt đầu bằng 1090."
            )
    
            mask_promo = pricelist.isin(["2010", "2030", "2050"])
    
            self._append_note_err(
            data,
            data.index[
                mask_promo
                & data["PROMOTION SALE PRICE"].fillna("").str.strip().eq("")
            ],
            "PROMOTION SALE PRICE không được để trống khi PRICELIST bắt đầu bằng 2010, 2030 hoặc 2050."
            )
            data["FILE NAME"] = path.name
            data["_SOURCE_ROW"] = data.index + 8
            data["STRUCTURE"] = self.dept[path.name]
            sources.append(data)

        self.src_listoff = pd.concat(sources, ignore_index=True)
    
        return self
        

    def _load_sitegroup(self) -> "Template_ETL":
        data = pd.read_excel(
            self.path_sitegroup,
            dtype=str,
            sheet_name="site-group",
        )

        data = (
            data
            .groupby("SITE_GROUP", as_index=False)
            .agg({"SITE": ";".join})
        )

        data["SITE_GROUP"] = data["SITE_GROUP"].astype(str).str.strip()
        data["SITE"] = data["SITE"].apply(self._sort_network)
        self.master_sitegroup_codes = set(data["SITE_GROUP"])
        self.sitegroup_members = {
            code: self._unique_sorted_sites(sites)
            for code, sites in zip(data["SITE_GROUP"], data["SITE"])
        }
        exact_matches = data.drop_duplicates(subset=["SITE"])
        self.sitegroup = dict(zip(exact_matches["SITE"], exact_matches["SITE_GROUP"]))

        activated = pd.read_excel(
            self.path_sitegroup,
            dtype=str,
            sheet_name="site-group-activated",
        )
        self._check_required_columns(activated, ["SITE_GROUP"])
        self.activated_sitegroup_codes = {
            str(code).strip()
            for code in activated["SITE_GROUP"].fillna("")
            if str(code).strip()
        }
        return self

    @staticmethod
    def _generate_sitegroup_code(unavailable_codes: set[str]) -> str:
        """Generate an available five-digit Site Group code."""
        first_offset = secrets.randbelow(90000)
        for offset in range(90000):
            code = str(10000 + ((first_offset + offset) % 90000))
            if code not in unavailable_codes:
                return code
        raise ValueError("Không còn mã SITE GROUP 5 chữ số khả dụng.")

    def _load_plan(self) -> "Template_ETL":
        plan = pd.read_excel(
            self.path_plan,
            dtype=str,
            header=3,
            sheet_name="plan-goldpromo",
        )

        self._check_required_columns(plan, date_columns_plan + ["CATALOGUE", "CATALOGUE DESCRIPTION"])
        self._check_required_data(plan, date_columns_plan + ["CATALOGUE", "CATALOGUE DESCRIPTION"])

        plan = plan.loc[plan["CATALOGUE"]==self.cata]
        
        plan[date_columns_plan] = plan[date_columns_plan].apply(
            pd.to_datetime,
            errors="coerce",
            dayfirst=True,
            format="mixed",
        )
        if plan[date_columns_plan].isna().any(axis=1).any():
            raise ValueError(f"CATALOGUE {self.cata} có ngày không hợp lệ trong Master data.")

        plan["SHOP ACTIVATION"] = (
            plan["CATALOGUE START DATE"] - pd.Timedelta(days=23)
        )
        plan["GLOBAL PERIOD START"] = plan["CATALOGUE START DATE"] - pd.Timedelta(days=24)
        plan["GLOBAL PERIOD END"] = plan["CATALOGUE END DATE"]

        plan["COMMITMENT CLOSING"] = plan["GENERAL PO DATE (D-17)"]
        plan["COMMITMENT DEADLINE"] = plan["COMMITMENT CLOSING"] - pd.Timedelta(days=1)

        plan["ORDER WAREHOUSE START"] = plan["CATALOGUE START DATE"] - pd.Timedelta(days=14)
        plan["ORDER WAREHOUSE END"] = plan["CATALOGUE END DATE"]

        self.plan = plan

        return self

    @staticmethod
    def _expand(value: str, network_dict: Optional[dict]) -> list:
        if value is None:
            return []

        value = str(value).strip()

        if value == "":
            return []

        if value.startswith("(") and value.endswith(")"):
            value = value[1:-1]

        network_dict = network_dict or {}

        def expand_token(token: str, parents: frozenset[str]) -> list[str]:
            token = token.strip()
            if not token:
                return []
            if token not in network_dict or token in parents:
                return [token]

            mapped = network_dict[token]
            mapped_tokens = "" if mapped is None else str(mapped)
            next_parents = parents | {token}
            expanded = []
            for mapped_token in mapped_tokens.split(";"):
                expanded.extend(expand_token(mapped_token, next_parents))
            return expanded

        result = []
        for token in value.split(";"):
            result.extend(expand_token(token, frozenset()))

        return result

    def _extract_discount_network(self, expression: str) -> str:
        return self._sort_network(self._extract_network(expression, network_key="DISCOUNT_NETWORK"))

    def _extract_network(
        self, expression: str, deduplicate: bool = True, network_key: str = "network"
    ) -> str:
        if pd.isna(expression) or expression is None:
            return ""

        expression = str(expression).strip().replace(" ", "")

        if expression == "":
            return ""

        # Evaluate each top-level network independently. Semicolons inside
        # parentheses belong to that network's adjustment expression.
        components = []
        start = 0
        depth = 0
        for position, character in enumerate(expression):
            if character == "(":
                depth += 1
            elif character == ")":
                depth = max(0, depth - 1)
            elif character == ";" and depth == 0:
                components.append(expression[start:position])
                start = position + 1
        components.append(expression[start:])
        components = [component for component in components if component]

        # A compact adjustment at the end applies to the union of the
        # preceding top-level networks. For example,
        # ``8210;8220(-112;126)`` means ``8210+8220-(112;126)``.
        if len(components) > 1:
            for compact_index, component in enumerate(components):
                compact = re.fullmatch(
                    r"([A-Za-z0-9]+)\(([^()]*)\)((?:[+-].*)?)",
                    component,
                )
                other_components = [
                    *components[:compact_index],
                    *components[compact_index + 1:],
                ]
                if compact is None or not all(
                    re.fullmatch(r"[A-Za-z0-9]+", value)
                    for value in other_components
                ):
                    continue

                adjustments = re.findall(r"[+-][^+-]+", compact.group(2))
                if not adjustments or "".join(adjustments) != compact.group(2):
                    continue

                expression = "+".join(
                    [*components[:compact_index], compact.group(1)]
                )
                for adjustment in adjustments:
                    sign = adjustment[0]
                    value = adjustment[1:]
                    expression += (
                        f"{sign}({value})"
                        if ";" in value
                        else f"{sign}{value}"
                    )
                expression += compact.group(3)
                for value in components[compact_index + 1:]:
                    expression += f"+{value}"
                components = [expression]
                break

        if len(components) > 1:
            result = []
            seen = set()
            for component in components:
                for site in self._parse_sites(
                    self._extract_network(component, deduplicate=deduplicate, network_key=network_key)
                ):
                    if not deduplicate or site not in seen:
                        seen.add(site)
                        result.append(site)
            return ";".join(result)

        # Support a base network with adjustments in parentheses followed by
        # standalone stores, e.g. ``8230(-132;135);112``.  Stores after the
        # closing parenthesis are additions to the adjusted base network.
        m = re.fullmatch(r'([A-Za-z0-9]+)\(([^()]*)\)(.*)', expression)

        if m:
            base = m.group(1)
            inside = m.group(2)
            suffix = m.group(3).strip(";")

            parts = re.findall(r'[+-][^+-]+', inside)

            if parts:
                expr = base

                for p in parts:
                    sign = p[0]
                    value = p[1:]

                    if ";" in value:
                        expr += f"{sign}({value})"
                    else:
                        expr += f"{sign}{value}"

                if suffix:
                    if suffix.startswith(("+", "-")):
                        expr += suffix
                    else:
                        expr += f"+({suffix})" if ";" in suffix else f"+{suffix}"

                expression = expr

        if not expression:
            return ""

        if "+" not in expression and "-" not in expression:
            result = []
            seen = set()

            for site in self._expand(expression, self.dict_network.get(network_key)):
                if not deduplicate or site not in seen:
                    seen.add(site)
                    result.append(site)

            return ";".join(result)

        m = re.match(r'([^+-]+)', expression)
        base_token = m.group(1) if m else ""

        result = []
        seen = set()

        for site in self._expand(base_token, self.dict_network.get(network_key)):
            if not deduplicate or site not in seen:
                seen.add(site)
                result.append(site)

        remain = expression[m.end():] if m else expression

        pattern = r'([+-])(\([^)]+\)|[^+-]+)'

        for op, value in re.findall(pattern, remain):
            sites = self._expand(value, self.dict_network.get(network_key))

            if op == "+":
                for s in sites:
                    if not deduplicate or s not in seen:
                        seen.add(s)
                        result.append(s)
            else:
                remove = set(sites)
                result = [x for x in result if x not in remove]
                seen = set(result)

        return ";".join(result)

    @staticmethod
    def _sort_key(site: str):
        site = "" if site is None else str(site)
        return (0, int(site)) if site.isdigit() else (1, site)

    @staticmethod
    def _parse_sites(text) -> list:
        if pd.isna(text):
            return []

        text = str(text).strip()

        if text == "":
            return []

        return [x.strip() for x in text.split(";") if x.strip()]

    def _sort_network(self, text) -> str:
        return ";".join(self._unique_sorted_sites(text))

    @staticmethod
    def _normalize_network_punctuation(text, replacement: str = ";") -> str:
        if pd.isna(text):
            return ""

        text = str(text)
        text = re.sub(r"[^\w\s+\-()]", replacement, text)
        text = re.sub(f"{re.escape(replacement)}+", replacement, text)
        return text

    @staticmethod
    def _has_valid_network_rule(expression) -> bool:
        """Return whether an expression follows the supported SITE RULE syntax."""
        if pd.isna(expression):
            return False

        expression = str(expression).strip().replace(" ", "")
        if not expression:
            return False

        components = []
        start = 0
        depth = 0
        for position, character in enumerate(expression):
            if character == "(":
                depth += 1
                if depth > 1:
                    return False
            elif character == ")":
                depth -= 1
                if depth < 0:
                    return False
            elif character == ";" and depth == 0:
                component = expression[start:position]
                if not component:
                    return False
                components.append(component)
                start = position + 1
        if depth != 0 or start == len(expression):
            return False
        components.append(expression[start:])

        token = r"[A-Za-z0-9]+"
        token_list = rf"{token}(?:;{token})*"
        standard = re.compile(rf"{token}(?:[+-](?:{token}|\({token_list}\)))+")
        plain = re.compile(rf"{token}|\({token_list}\)")
        compact = re.compile(
            rf"({token})\(([^()]*)\)((?:[+-](?:{token}|\({token_list}\)))*)"
        )

        for component in components:
            if plain.fullmatch(component) or standard.fullmatch(component):
                continue

            match = compact.fullmatch(component)
            if match is None:
                return False
            inside = match.group(2)
            adjustments = re.findall(r"[+-][^+-]+", inside)
            if not adjustments or "".join(adjustments) != inside:
                return False
            for index, adjustment in enumerate(adjustments):
                raw_values = adjustment[1:]
                if raw_values.startswith(";") or (
                    raw_values.endswith(";") and index == len(adjustments) - 1
                ):
                    return False
                values = raw_values.rstrip(";")
                if not values or re.fullmatch(token_list, values) is None:
                    return False

        return True

    def _check_network(self, data) -> Optional[pd.DataFrame]:
        data = self._ensure_note_err(data)
        # This marker prevents dependent validation from treating a blocked
        # network expression as an empty or mismatched expanded network.
        data["_SKIP_NETWORK_CHECKS"] = False

        data["PURCHASE NETWORK"] = data["PURCHASE NETWORK"].map(self._normalize_network_punctuation)
        data["GOLD PROMO NETWORK"] = data["GOLD PROMO NETWORK"].map(self._normalize_network_punctuation)

        for _, idx in data.groupby(self.GROUP_COLS, dropna=False).groups.items():
            rows = data.loc[idx]
            uses_network_8000 = rows[
                ["PURCHASE NETWORK", "GOLD PROMO NETWORK"]
            ].fillna("").astype(str).apply(
                lambda column: column.str.strip().eq("8000")
            ).any().any()
            if uses_network_8000:
                data.loc[idx, "_SKIP_NETWORK_CHECKS"] = True
                self._append_note_err(data, idx, "Không được sử dụng network 8000.")

        active_rows = ~data["_SKIP_NETWORK_CHECKS"]
        data["PURCHASE NETWORK EXPANDED"] = ""
        data["DISCOUNT_NETWORK_EXPANDED"] = ""
        data["GOLD PROMO NETWORK EXPANDED"] = ""
        data.loc[active_rows, "PURCHASE NETWORK EXPANDED"] = data.loc[
            active_rows, "PURCHASE NETWORK"
        ].map(self._extract_network)
        data.loc[active_rows, "DISCOUNT_NETWORK_EXPANDED"] = data.loc[
            active_rows, "PURCHASE NETWORK"
        ].map(self._extract_discount_network)
        data.loc[active_rows, "GOLD PROMO NETWORK EXPANDED"] = data.loc[
            active_rows, "GOLD PROMO NETWORK"
        ].map(self._extract_network)

        valid_stores = set(self.dict_network.get("store_hyper", []))

        for _, idx in data.groupby(self.GROUP_COLS, dropna=False).groups.items():
            rows = data.loc[idx]
            if rows["_SKIP_NETWORK_CHECKS"].any():
                continue

            messages = []
            invalid_site_rule = False
            for source_column, expanded_column in (
                ("PURCHASE NETWORK", "PURCHASE NETWORK EXPANDED"),
                ("GOLD PROMO NETWORK", "GOLD PROMO NETWORK EXPANDED"),
            ):
                valid_rule = rows[source_column].map(self._has_valid_network_rule)
                if (~valid_rule).any():
                    messages.append(f"{source_column} không đúng SITE RULE.")
                    invalid_site_rule = True

            if invalid_site_rule:
                data.loc[idx, "_SKIP_NETWORK_CHECKS"] = True
                self._append_note_err(data, idx, " | ".join(messages))
                continue

            for source_column, expanded_column in (
                ("PURCHASE NETWORK", "PURCHASE NETWORK EXPANDED"),
                ("GOLD PROMO NETWORK", "GOLD PROMO NETWORK EXPANDED"),
            ):
                valid_rule = rows[source_column].map(self._has_valid_network_rule)
                has_adjustment = rows[source_column].astype(str).str.contains(
                    r"[+-]", regex=True, na=False
                )
                empty_after_expand = (
                    valid_rule
                    & has_adjustment
                    & rows[expanded_column].fillna("").astype(str).str.strip().eq("")
                )
                if empty_after_expand.any():
                    messages.append(
                        f"{source_column} sau khi expand +/- bị rỗng, vui lòng kiểm tra lại."
                    )

            invalid_sites = set()
            for col in ["PURCHASE NETWORK EXPANDED", "GOLD PROMO NETWORK EXPANDED"]:
                for value in rows[col]:
                    invalid_sites.update(
                        site for site in self._parse_sites(value)
                        if site not in valid_stores
                    )

            if invalid_sites:
                invalid_sorted = sorted(invalid_sites, key=self._sort_key)
                messages.append(
                    "Cửa hàng " + ";".join(invalid_sorted)
                    + " không thuộc network 8200 hiện tại"
                )

            promo_variants = {
                self._unique_sorted_sites(value)
                for value in rows["GOLD PROMO NETWORK EXPANDED"]
            }
            if len(promo_variants) > 1:
                messages.append("Vui lòng kiểm tra lại GOLD PROMO NETWORK")

            self._append_note_err(data, idx, " | ".join(messages))

        return data

    def _ppNetwork_gpNetwork(self, data) -> Optional[pd.DataFrame]:
        data = self._ensure_note_err(data)

        for _, related_indices in data.groupby(
            ["GOLD CODE", "LV"], dropna=False
        ).groups.items():
            related_indices = pd.Index(related_indices)
            if "_SKIP_NETWORK_CHECKS" in data.columns:
                related_indices = related_indices[
                    ~data.loc[related_indices, "_SKIP_NETWORK_CHECKS"]
                ]
            if related_indices.empty:
                continue
            duplicates = set()
            for value in data.loc[related_indices, "PURCHASE NETWORK"]:
                expanded_with_duplicates = self._parse_sites(
                    self._extract_network(value, deduplicate=False)
                )
                counts = Counter(expanded_with_duplicates)
                duplicates.update(
                    site for site, count in counts.items() if count > 1
                )
            if duplicates:
                sorted_duplicates = sorted(duplicates, key=self._sort_key)
                self._append_note_err(
                    data,
                    pd.Index(related_indices),
                    "PURCHASE NETWORK bị trùng lặp: "
                    + ";".join(sorted_duplicates),
                )

        for _, idx in data.groupby(self.GROUP_COLS, dropna=False).groups.items():
            rows = data.loc[idx]
            if rows.get("_SKIP_NETWORK_CHECKS", pd.Series(False, index=rows.index)).any():
                continue

            purchase_lists = [
                self._unique_sorted_sites(value)
                for value in rows["PURCHASE NETWORK EXPANDED"]
            ]

            messages = []

            if len(set(purchase_lists)) > 1:
                counter = Counter()
                for sites in purchase_lists:
                    counter.update(sites)

                dup = sorted(
                    (site for site, cnt in counter.items() if cnt > 1),
                    key=self._sort_key,
                )

                if dup:
                    messages.append("PURCHASE NETWORK bị trùng lặp: " + ";".join(dup))

            purchase = set().union(*purchase_lists) if purchase_lists else set()

            promo = set()
            for value in rows["GOLD PROMO NETWORK EXPANDED"]:
                promo.update(self._parse_sites(value))

            missing = sorted(promo - purchase, key=self._sort_key)
            extra = sorted(purchase - promo, key=self._sort_key)

            if missing:
                messages.append("Thiếu cửa hàng trong PURCHASE NETWORK: " + ";".join(missing))
            if extra:
                messages.append("Dư cửa hàng trong PURCHASE NETWORK: " + ";".join(extra))

            self._append_note_err(data, idx, " | ".join(messages))

        for col in ["PURCHASE NETWORK EXPANDED", "GOLD PROMO NETWORK EXPANDED"]:
            data[col] = data[col].map(self._sort_network)

        return data

    def _get_sitegroup(self, data) -> Optional[pd.DataFrame]:
        """Populate only exact master Site Group matches.

        Values in the source ``SITE GROUP`` column are deliberately ignored
        here.  A Site Group must always be determined from the expanded Gold
        Promo store list and the master Site Group member lists.
        """
        data["SITE GROUP"] = data["GOLD PROMO NETWORK EXPANDED"].map(self.sitegroup).fillna("")
        return data

    def get_sitegroup_suggestions(self) -> list[dict]:
        """Generate a new five-digit code for every non-exact Site Group."""
        if self.src is None:
            return []

        data = self.src
        suggestions = []
        # Exact master matches have already been written to ``SITE GROUP`` by
        # _get_sitegroup.  Those codes are no longer candidates for a new
        # GOLD PROMO NETWORK EXPANDED value.
        assigned_codes = {
            str(code).strip()
            for code in data["SITE GROUP"]
            if pd.notna(code) and str(code).strip()
        }
        suggested_codes: set[str] = set()
        for network, rows in data.groupby("GOLD PROMO NETWORK EXPANDED", sort=True):
            network = "" if pd.isna(network) else str(network)
            # Exact matches have already been resolved from the master list.
            if str(rows["SITE GROUP"].iat[0]).strip():
                continue

            unavailable_new_codes = (
                self.master_sitegroup_codes
                | self.activated_sitegroup_codes
                | assigned_codes
                | suggested_codes
                | self.non_suggested_sitegroup_codes
            )
            code = self._generate_sitegroup_code(unavailable_new_codes)
            suggested_codes.add(code)
            raw_networks = sorted(
                {str(value) for value in rows["GOLD PROMO NETWORK"].dropna() if str(value).strip()}
            )
            structures = sorted(
                {
                    f"0{str(value).strip()}"
                    for value in rows["STRUCTURE"].dropna()
                    if str(value).strip()
                },
                key=self._sort_key,
            )
            suggestions.append(
                {
                    "structure": ";".join(structures),
                    "gold_promo_network": "; ".join(raw_networks),
                    "expanded_network": network,
                    "suggested_code": code,
                    "original_suggested_code": "",
                }
            )

        return suggestions

    def validate_sitegroup_changes(self, suggestions: list[dict]) -> list[str]:
        """Return Site Group codes that would duplicate the master file."""
        duplicate_codes: set[str] = set()
        new_codes: set[str] = set()
        for suggestion in suggestions:
            original = str(suggestion.get("original_suggested_code", "")).strip()
            selected = str(suggestion.get("suggested_code", "")).strip()
            if selected == original or not selected:
                continue
            if (
                selected in self.master_sitegroup_codes
                or selected in self.activated_sitegroup_codes
                or selected in new_codes
            ):
                duplicate_codes.add(selected)
            new_codes.add(selected)
        return sorted(duplicate_codes, key=self._sort_key)

    # def update_sitegroup_file(self, suggestions: list[dict]) -> bool:
    #     """Apply confirmed Site Group additions/removals to the master sheet."""
    #     duplicate_codes = self.validate_sitegroup_changes(suggestions)
    #     if duplicate_codes:
    #         raise ValueError("Duplicate Site Group: " + "; ".join(duplicate_codes))

    #     delete_codes: set[str] = set()
    #     new_rows: list[tuple[str, str]] = []
    #     replacement_rows: dict[str, tuple[str, ...]] = {}
    #     for suggestion in suggestions:
    #         original = str(suggestion.get("original_suggested_code", "")).strip()
    #         selected = str(suggestion.get("suggested_code", "")).strip()
    #         if selected == original:
    #             if original:
    #                 target_sites = self._unique_sorted_sites(suggestion["expanded_network"])
    #                 if set(self.sitegroup_members.get(original, ())) != set(target_sites):
    #                     replacement_rows[original] = target_sites
    #             continue
    #         if original:
    #             delete_codes.add(original)
    #         if selected:
    #             new_rows.append((selected, str(suggestion["expanded_network"]).strip()))

    #     if not delete_codes and not new_rows and not replacement_rows:
    #         return False

    #     keep_vba = self.path_sitegroup.suffix.lower() == ".xlsm"
    #     workbook = load_workbook(self.path_sitegroup, keep_vba=keep_vba)
    #     temporary_path = None
    #     try:
    #         sheet = workbook["site-group"]
    #         headers = {
    #             str(sheet.cell(1, column).value).strip(): column
    #             for column in range(1, sheet.max_column + 1)
    #         }
    #         code_column = headers.get("SITE_GROUP")
    #         site_column = headers.get("SITE")
    #         if code_column is None or site_column is None:
    #             raise ValueError("Sheet site-group must contain SITE_GROUP and SITE columns.")

    #         for row in range(sheet.max_row, 1, -1):
    #             value = sheet.cell(row, code_column).value
    #             if value is not None and str(value).strip() in (delete_codes | set(replacement_rows)):
    #                 sheet.delete_rows(row, 1)

    #         rows_to_add = [
    #             *replacement_rows.items(),
    #             *((code, self._unique_sorted_sites(sites)) for code, sites in new_rows),
    #         ]
    #         for code, sites in rows_to_add:
    #             for site in sites:
    #                 sheet.append([
    #                     code if column == code_column else site if column == site_column else None
    #                     for column in range(1, sheet.max_column + 1)
    #                 ])

    #         with tempfile.NamedTemporaryFile(
    #             suffix=self.path_sitegroup.suffix,
    #             dir=self.path_sitegroup.parent,
    #             delete=False,
    #         ) as temporary_file:
    #             temporary_path = Path(temporary_file.name)
    #         workbook.save(temporary_path)
    #         workbook.close()
    #         os.replace(temporary_path, self.path_sitegroup)
    #     except Exception:
    #         workbook.close()
    #         if temporary_path is not None and temporary_path.exists():
    #             temporary_path.unlink()
    #         raise

    #     self.master_sitegroup_codes.difference_update(delete_codes)
    #     self.master_sitegroup_codes.update(code for code, _ in new_rows)
    #     for code in delete_codes:
    #         self.sitegroup_members.pop(code, None)
    #     self.sitegroup_members.update(replacement_rows)
    #     self.sitegroup_members.update(
    #         {code: self._unique_sorted_sites(sites) for code, sites in new_rows}
    #     )
    #     return True

    def update_sitegroup_file(self, suggestions: list[dict]) -> bool:
        with workbook_write_lock(self.path_sitegroup):
            return self._update_sitegroup_file_locked(suggestions)

    def _update_sitegroup_file_locked(self, suggestions: list[dict]) -> bool:
            """Apply confirmed Site Group additions/removals to the master sheet."""
            duplicate_codes = self.validate_sitegroup_changes(suggestions)
            if duplicate_codes:
                raise ValueError("Duplicate Site Group: " + "; ".join(duplicate_codes))

            delete_codes: set[str] = set()
            new_rows: list[tuple[str, list[str]]] = []
            replacement_rows: dict[str, list[str]] = {}
            for suggestion in suggestions:
                original = str(suggestion.get("original_suggested_code", "")).strip()
                selected = str(suggestion.get("suggested_code", "")).strip()
                if selected == original:
                    if original:
                        target_sites = self._unique_sorted_sites(suggestion["expanded_network"])
                        if set(self.sitegroup_members.get(original, ())) != set(target_sites):
                            replacement_rows[original] = target_sites
                    continue
                if original:
                    delete_codes.add(original)
                if selected:
                    # Chuẩn hóa về list[str] ngay tại đây, tránh gọi lại _unique_sorted_sites
                    # trên một string (dễ bug: iterate theo ký tự) ở bước sau.
                    new_rows.append((selected, self._unique_sorted_sites(suggestion["expanded_network"])))

            codes_to_delete = delete_codes | set(replacement_rows)
            if not codes_to_delete and not new_rows and not replacement_rows:
                return False

            keep_vba = self.path_sitegroup.suffix.lower() == ".xlsm"
            workbook = load_workbook(self.path_sitegroup, keep_vba=keep_vba)
            temporary_path = None
            try:
                sheet = workbook["site-group"]
                activated_sheet = workbook["site-group-activated"]
                headers = {
                    str(sheet.cell(1, column).value).strip(): column
                    for column in range(1, sheet.max_column + 1)
                }
                code_column = headers.get("SITE_GROUP")
                site_column = headers.get("SITE")
                if code_column is None or site_column is None:
                    raise ValueError("Sheet site-group must contain SITE_GROUP and SITE columns.")

                activated_code_column = next(
                    (
                        column
                        for column in range(1, activated_sheet.max_column + 1)
                        if str(activated_sheet.cell(1, column).value).strip()
                        == "SITE_GROUP"
                    ),
                    None,
                )
                if activated_code_column is None:
                    raise ValueError(
                        "Sheet site-group-activated must contain the SITE_GROUP column."
                    )
                activated_codes = {
                    str(activated_sheet.cell(row, activated_code_column).value).strip()
                    for row in range(2, activated_sheet.max_row + 1)
                    if activated_sheet.cell(row, activated_code_column).value is not None
                    and str(
                        activated_sheet.cell(row, activated_code_column).value
                    ).strip()
                }
                activated_duplicates = sorted(
                    {
                        code
                        for code, _ in new_rows
                        if code in activated_codes
                    },
                    key=self._sort_key,
                )
                if activated_duplicates:
                    raise ValueError(
                        "Duplicate Site Group in site-group-activated: "
                        + "; ".join(activated_duplicates)
                    )

                max_row = sheet.max_row
                max_col = sheet.max_column

                # Đọc toàn bộ dữ liệu 1 lần (nhanh hơn nhiều so với .cell() từng ô),
                # lọc bỏ các dòng cần xóa trong bộ nhớ.
                kept_rows: list[tuple] = []
                if max_row > 1:
                    for row in sheet.iter_rows(min_row=2, max_row=max_row, values_only=True):
                        code_val = row[code_column - 1]
                        if code_val is None or str(code_val).strip() not in codes_to_delete:
                            kept_rows.append(row)

                    # Xóa toàn bộ vùng dữ liệu (row 2..max_row) trong MỘT lần gọi,
                    # thay vì delete_rows từng dòng riêng lẻ.
                    sheet.delete_rows(2, max_row - 1)

                # Ghi lại các dòng giữ nguyên
                for row in kept_rows:
                    sheet.append(list(row))

                # Ghi các dòng mới/thay thế, dùng template list dựng sẵn thay vì
                # list comprehension quét toàn bộ số cột cho mỗi site.
                rows_to_add = [*replacement_rows.items(), *new_rows]
                template = [None] * max_col
                for code, sites in rows_to_add:
                    for site in sites:
                        new_row = template.copy()
                        new_row[code_column - 1] = code
                        new_row[site_column - 1] = site
                        sheet.append(new_row)

                # Only newly created codes are appended to the activated list;
                # replacements of existing Site Groups do not add another row.
                activated_template = [None] * activated_sheet.max_column
                for code, _ in new_rows:
                    activated_row = activated_template.copy()
                    activated_row[activated_code_column - 1] = code
                    activated_sheet.append(activated_row)

                with tempfile.NamedTemporaryFile(
                    suffix=self.path_sitegroup.suffix,
                    dir=self.path_sitegroup.parent,
                    delete=False,
                ) as temporary_file:
                    temporary_path = Path(temporary_file.name)
                workbook.save(temporary_path)
                workbook.close()
                os.replace(temporary_path, self.path_sitegroup)
            except Exception:
                workbook.close()
                if temporary_path is not None and temporary_path.exists():
                    temporary_path.unlink()
                raise

            self.master_sitegroup_codes.difference_update(delete_codes)
            self.master_sitegroup_codes.update(code for code, _ in new_rows)
            for code in delete_codes:
                self.sitegroup_members.pop(code, None)
            self.sitegroup_members.update(replacement_rows)
            self.sitegroup_members.update(dict(new_rows))
            self.activated_sitegroup_codes.update(code for code, _ in new_rows)
            return True

    def record_sitegroup_history(self) -> int:
        """Record every resolved Site Group from every source in this run."""
        if self.check_attribute:
            return 0
        if self.src is None or self.path_plan is None or not str(self.cata).strip():
            raise ValueError("Source, master workbook and catalogue are required for SITE_GROUP_CHECK.")
        columns = ["SITE GROUP", "GOLD PROMO NETWORK EXPANDED"]
        self._check_required_columns(self.src, columns)
        rows = set()
        for group, network in self.src[columns].itertuples(index=False, name=None):
            code = "" if pd.isna(group) else str(group).strip()
            sites = self._unique_sorted_sites(network)
            if not code or not sites:
                raise ValueError("SITE_GROUP_CHECK requires a Site Group and sites for every source row.")
            rows.update((str(self.cata).strip(), code, site) for site in sites)
        return append_sitegroup_history(self.path_plan, rows)

    def apply_sitegroup_suggestions(self, suggestions: list[dict]) -> list[dict]:
        """Apply Site Group codes confirmed or entered by the user."""
        if self.src is None:
            return []

        for suggestion in suggestions:
            network = suggestion["expanded_network"]
            code = str(suggestion["suggested_code"]).strip()
            if not code:
                continue
            self.src.loc[self.src["GOLD PROMO NETWORK EXPANDED"].eq(network), "SITE GROUP"] = code
        return []

    def _getSO(self, data) -> Optional[pd.DataFrame]:
        data["SO"] = self.cata + "D" + data["STRUCTURE"].astype(str)
        codes = pd.factorize(data["GOLD PROMO NETWORK EXPANDED"])[0] + 1
        data["ID_SO"] = [f"{x:02d}" for x in codes]
        data["SO"] = data["SO"] + "-" + data["ID_SO"]
        data = data.drop(columns=["ID_SO"])

        return data

    def build_gold_promo_network_report(self) -> pd.DataFrame:
        """Summarize unique articles and resulting OA rows for every SO.

        The report is built from a copy so creating it cannot modify the
        validated source data used by the rest of the Stage 1 workflow.
        """
        if self.src is None:
            raise ValueError("Stage 1 source data has not been loaded.")

        required_columns = {
            "SO",
            "GOLD PROMO NETWORK EXPANDED",
            "GOLD CODE",
            "LV",
            "LU",
        }
        missing_columns = required_columns.difference(self.src.columns)
        if missing_columns:
            missing = ", ".join(sorted(missing_columns))
            raise ValueError(f"Cannot create network report; missing columns: {missing}")

        report_rows = self.src[
            ["SO", "GOLD PROMO NETWORK EXPANDED", "GOLD CODE", "LV", "LU"]
        ].drop_duplicates().copy()
        report_rows["SO"] = report_rows["SO"].fillna("").astype(str).str.strip()
        report_rows["GOLD PROMO NETWORK EXPANDED"] = (
            report_rows["GOLD PROMO NETWORK EXPANDED"]
            .fillna("")
            .astype(str)
            .str.strip()
        )

        report = (
            report_rows.groupby(
                ["SO", "GOLD PROMO NETWORK EXPANDED"],
                sort=False,
                dropna=False,
            )
            .size()
            .rename("GOLD CODE + LV + LU COUNT")
            .reset_index()
        )
        network_counts = report["GOLD PROMO NETWORK EXPANDED"].map(
            lambda network: sum(bool(site.strip()) for site in network.split(";"))
        )
        report["OA COUNT"] = report["GOLD CODE + LV + LU COUNT"] * network_counts

        total = pd.DataFrame(
            [{
                "SO": "TOTAL",
                "GOLD PROMO NETWORK EXPANDED": "",
                "GOLD CODE + LV + LU COUNT": int(
                    report["GOLD CODE + LV + LU COUNT"].sum()
                ),
                "OA COUNT": int(report["OA COUNT"].sum()),
            }]
        )
        return pd.concat([report, total], ignore_index=True)

    def _check_allocation(self, data) -> Optional[pd.DataFrame]:
        delivery_columns = ["% DELIVERY 1", "% DELIVERY 2", "% DELIVERY 3"]
        site_columns = [
            col for col in data.columns
            if col in self.dict_network.get("store", [])
        ]

        group_keys = [
            "GOLD CODE",
            "LV",
            "LU",
        ]

        data = self._ensure_note_err(data)

        # Accept values such as ``50`` and ``50%``. Blank delivery slots are
        # treated as zero, so any populated subset must still total 100.
        delivery_text = data[delivery_columns].fillna("").astype(str).apply(
            lambda column: column.str.strip().str.replace("%", "", regex=False).str.strip()
        )
        delivery_values = delivery_text.apply(pd.to_numeric, errors="coerce")
        invalid_delivery = delivery_text.ne("") & delivery_values.isna()
        invalid_rows = invalid_delivery.any(axis=1)
        self._append_note_err(
            data,
            data.index[invalid_rows],
            "% DELIVERY 1, % DELIVERY 2 và % DELIVERY 3 phải là số.",
        )

        delivery_total = delivery_values.fillna(0).sum(axis=1)
        invalid_total = ~invalid_rows & ~delivery_total.round(10).eq(100)
        self._append_note_err(
            data,
            data.index[invalid_total],
            "Tổng % DELIVERY 1 + % DELIVERY 2 + % DELIVERY 3 phải bằng 100.",
        )

        for _, idx in data.groupby(group_keys, dropna=False).groups.items():
            rows = data.loc[idx]

            missing_allocation_sites = set()
            invalid_allocation_sites = set()
            for _, row in rows.iterrows():
                row_purchase_sites = {
                    site for site in self._parse_sites(row["PURCHASE NETWORK EXPANDED"])
                    if site in site_columns
                }
                for site in row_purchase_sites:
                    value = row[site]
                    if pd.isna(value) or str(value).strip() == "":
                        missing_allocation_sites.add(site)
                        continue
                    numeric_value = pd.to_numeric(pd.Series([value]), errors="coerce").iat[0]
                    if (
                        pd.isna(numeric_value)
                        or numeric_value <= 0
                        or numeric_value % 1 != 0
                    ):
                        invalid_allocation_sites.add(site)

            missing_sites = sorted(missing_allocation_sites, key=self._sort_key)
            invalid_allocation_sites = sorted(
                invalid_allocation_sites,
                key=self._sort_key,
            )

            if invalid_allocation_sites:
                self._append_note_err(
                    data,
                    idx,
                    "Thông tin phân bổ phải là số nguyên và lớn hơn 0: "
                    + ", ".join(invalid_allocation_sites)
                    + ".",
                )

            if not missing_sites:
                continue

            message = (
                f"Thiếu phân bổ đối với các cửa hàng: {';'.join(missing_sites)} "
                "dựa trên PURCHASE NETWORK EXPANDED."
            )

            self._append_note_err(data, idx, message)

        return data

    def _validate_structure_gold_lv(self, data: pd.DataFrame) -> pd.DataFrame:
        """Flag a GOLD CODE/LV combination assigned to multiple structures."""
        self._ensure_note_err(data)
        keys = ["GOLD CODE", "LV"]
        normalized = data.loc[:, [*keys, "STRUCTURE"]].fillna("").astype(str).apply(
            lambda column: column.str.strip()
        )
        valid_keys = normalized["GOLD CODE"].ne("") & normalized["LV"].ne("")

        for _, index in normalized.loc[valid_keys].groupby(keys, dropna=False).groups.items():
            structures = set(normalized.loc[index, "STRUCTURE"]) - {""}
            if len(structures) > 1:
                self._append_note_err(
                    data,
                    pd.Index(index),
                    "GOLD CODE và LV trùng nhau nhưng STRUCTURE khác nhau.",
                )
        return data

    def _validate_overlapping_price_or_discount(
        self,
        data: pd.DataFrame,
    ) -> pd.DataFrame:
        """Flag conflicting prices or discounts for each expanded network."""
        self._ensure_note_err(data)
        key_columns = [
            "GOLD CODE",
            "LV",
        ]
        value_columns = [
            "NORMAL PURCHASE PRICE",
            "DISCOUNT (% OR VALUE)",
        ]

        comparison = data.loc[
            :, [*key_columns, "PURCHASE NETWORK EXPANDED", *value_columns]
        ].copy()
        comparison["_SOURCE_INDEX"] = data.index

        for column in [*key_columns, *value_columns]:
            comparison[column] = (
                comparison[column].fillna("").astype(str).str.strip()
            )

        comparison["PURCHASE NETWORK EXPANDED"] = comparison[
            "PURCHASE NETWORK EXPANDED"
        ].map(self._parse_sites)
        comparison = comparison.explode("PURCHASE NETWORK EXPANDED")
        comparison["PURCHASE NETWORK EXPANDED"] = comparison[
            "PURCHASE NETWORK EXPANDED"
        ].fillna("").astype(str).str.strip()
        comparison = comparison[
            comparison["PURCHASE NETWORK EXPANDED"].ne("")
        ]

        group_columns = [*key_columns, "PURCHASE NETWORK EXPANDED"]
        conflicting_indices = set()
        for _, rows in comparison.groupby(group_columns, dropna=False):
            price_conflict = rows["NORMAL PURCHASE PRICE"].nunique(dropna=False) > 1
            discount_conflict = self._has_invalid_discount_difference(
                rows["DISCOUNT (% OR VALUE)"]
            )
            if price_conflict or discount_conflict:
                conflicting_indices.update(rows["_SOURCE_INDEX"])

        if conflicting_indices:
            self._append_note_err(
                data,
                data.index[data.index.isin(conflicting_indices)],
                "Thông tin mua hàng và chiết khấu bị trùng.",
            )

        return data

    @staticmethod
    def _discount_type(value: str) -> str:
        """Return the AG discount type used by the discount templates."""
        text = "" if pd.isna(value) else str(value).strip()
        if "+" in text:
            return "3"
        if "%" in text:
            return "1"
        return "2"

    @classmethod
    def _has_invalid_discount_difference(cls, discounts: pd.Series) -> bool:
        """Allow different discounts only between type 1/3 or type 2/3."""
        distinct = list(dict.fromkeys(discounts.fillna("").astype(str).str.strip()))
        for index, left in enumerate(distinct):
            for right in distinct[index + 1:]:
                type_pair = frozenset((cls._discount_type(left), cls._discount_type(right)))
                if type_pair not in cls.ALLOWED_DIFFERENT_DISCOUNT_TYPE_PAIRS:
                    return True
        return False

    def _validate_duplicate_purchase_information(
        self,
        data: pd.DataFrame,
    ) -> pd.DataFrame:
        """Flag rows with identical purchase and discount information."""
        self._ensure_note_err(data)
        duplicate_columns = [
            "GOLD CODE",
            "LV",
            "LU",
            "SUPPLIER CODE",
            "COMMERCIAL CONTRACT",
            "PURCHASE NETWORK",
            "NORMAL PURCHASE PRICE",
            "DISCOUNT (% OR VALUE)",
        ]
        normalized = data.loc[:, duplicate_columns].fillna("").astype(str).apply(
            lambda column: column.str.strip()
        )
        duplicate_rows = normalized.duplicated(
            subset=duplicate_columns,
            keep=False,
        )
        self._append_note_err(
            data,
            data.index[duplicate_rows],
            "Thông tin mua hàng và chiết khấu bị trùng.",
        )
        return data

    def _field_validator(self, data: pd.DataFrame) -> pd.DataFrame:
        lu_values = data["LU"].fillna("").astype(str).str.strip()
        self._append_note_err(
            data,
            data.index[~lu_values.isin(["1", "41"])],
            "LU chỉ được phép là 1 hoặc 41."
        )

        valid_delivery_type = {
            "CROSS-DOCKING",
            "DIRECT",
            "VINAMILK",
        }

        self._append_note_err(
            data,
            data.index[
                ~data["DELIVERY TYPE"].astype(str).str.strip().isin(valid_delivery_type)
            ],
            "DELIVERY TYPE chỉ được phép là CROSS-DOCKING, DIRECT hoặc VINAMILK."
        )

        return data

    def _contract_checking(self, x):
        if pd.isna(x):
            return ""

        contract = str(x).strip()
        if not contract:
            return ""

        if len(contract) == 8:
            if "YV000" in contract:
                return contract
            if contract[5:] in self.dict_network.get("wh", []):
                return contract
            return contract[:4]
        if any(s in contract for s in ("YV00", "YV0")):
            return f"{contract[:3]}YV000"
        return contract[:4]

    def _convert_date(self, data: Optional[pd.DataFrame]) -> Optional[pd.DataFrame]:
        data["PP START DATE"], data["PP END DATE"] = self._combine_date_columns(data, "PP")

        today = pd.Timestamp.today().normalize()

        mask = data["PP START DATE"].notna() & (data["PP START DATE"] <= today)
        data.loc[mask, "PP START DATE"] = today + pd.Timedelta(days=1)

        mask_err = (
            data["PP START DATE"].notna()
            & data["PP END DATE"].notna()
            & (data["PP START DATE"] > data["PP END DATE"])
        )

        self._append_note_err(
            data,
            data.index[mask_err],
            "PP START DATE phải nhỏ hơn hoặc bằng PP END DATE."
        )

        return data

    def _validate_pp_dates_against_plan(self, data: pd.DataFrame) -> pd.DataFrame:
        """Validate PP dates against order dates for the row's delivery type."""
        if self.plan is None or self.plan.empty:
            raise ValueError(f"Không tìm thấy CATALOGUE {self.cata} trong Master data.")

        self._ensure_note_err(data)
        delivery_types = ("CROSS-DOCKING", "DIRECT", "VINAMILK")
        normalized_delivery_type = (
            data["DELIVERY TYPE"].fillna("").astype(str).str.strip()
        )

        for delivery_type in delivery_types:
            row_mask = normalized_delivery_type.eq(delivery_type)
            order_date_1_column = f"{delivery_type} | ORDER DATE 1"
            order_date_3_column = f"{delivery_type} | ORDER DATE 3"
            order_date_1 = self.plan[order_date_1_column].iloc[0]
            order_date_3 = self.plan[order_date_3_column].iloc[0]

            invalid_start = (
                row_mask
                & data["PP START DATE"].notna()
                & data["PP START DATE"].ne(order_date_1)
            )
            self._append_note_err(
                data,
                data.index[invalid_start],
                f"PP START DATE phải bằng {order_date_1_column} "
                f"({order_date_1.strftime('%d/%m/%Y')}).",
            )

            invalid_end = (
                row_mask
                & data["PP END DATE"].notna()
                & data["PP END DATE"].lt(order_date_3)
            )
            self._append_note_err(
                data,
                data.index[invalid_end],
                f"PP END DATE phải lớn hơn hoặc bằng {order_date_3_column} "
                f"({order_date_3.strftime('%d/%m/%Y')}).",
            )

        return data

    def _pipeline(self) -> "Template_ETL":
        data = self.src

        data = self._check_network(data)
        data = self._ppNetwork_gpNetwork(data)
        source_structure = data["FILE NAME"].map(self.dept)
        blank_structure = data["STRUCTURE"].fillna("").astype(str).str.strip().eq("")
        data.loc[blank_structure, "STRUCTURE"] = source_structure.loc[blank_structure]
        if self.should_generate_so:
            # Validate always regenerates SO. SITE GROUP is populated later by
            # the separate Add Site Group action in the desktop workflow.
            data["STRUCTURE"] = source_structure
            data = self._getSO(data)
        data = self._field_validator(data)
        data = self._validate_structure_gold_lv(data)
        data = self._check_allocation(data)
        data = self._convert_date(data)
        data = self._validate_pp_dates_against_plan(data)

        data = self._validate_duplicate_purchase_information(data)
        data["COMMERCIAL CONTRACT"] = data["COMMERCIAL CONTRACT"].map(self._contract_checking)
        data = self._validate_overlapping_price_or_discount(data)

        site_columns = [
            col for col in data.columns
            if col in self.dict_network.get("store", [])
        ]

        mask = [
            *required_src,
            "NOTE ERR FROM MASTER DATA",
            "FILE NAME",
            "_SOURCE_ROW",
            "SITE GROUP",
            "SO",
            *site_columns,
            "PURCHASE NETWORK EXPANDED",
            "GOLD PROMO NETWORK EXPANDED",
            "DISCOUNT_NETWORK_EXPANDED",
            "PP START DATE",
            "PP END DATE"
        ]

        self.src = data[mask]
        self.non_warehouse_src = self._get_non_warehouse_src(
            self.src,
            self.exception_discount_gold_codes,
        )

        return self

    @classmethod
    def _get_non_warehouse_src(
        cls,
        data: pd.DataFrame,
        exception_discount_gold_codes=None,
    ) -> pd.DataFrame:
        """Return ``n+m`` discount rows that do not use T/TH notation.

        The returned frame is a copy, so the primary source frame remains
        unchanged and continues through every existing template flow.
        """
        discounts = (
            data["DISCOUNT (% OR VALUE)"]
            .fillna("")
            .astype(str)
            .str.replace(r"\s+", "", regex=True)
        )
        mask = discounts.str.fullmatch(cls.NON_WAREHOUSE_DISCOUNT, na=False)
        exception_codes = {
            str(code).strip()
            for code in (exception_discount_gold_codes or [])
            if str(code).strip()
        }
        if exception_codes and "GOLD CODE" in data.columns:
            gold_codes = data["GOLD CODE"].fillna("").astype(str).str.strip()
            mask &= ~gold_codes.isin(exception_codes)
        return data.loc[mask].copy()

    def _pipeline2(self) -> "Template_ETL":
        data = self.src_listoff

        data = data[required_stage2 + [
            'SP START DAY', 'SP START MONTH', 'SP START YEAR',
            'SP END DAY', 'SP END MONTH', 'SP END YEAR',
            'NOTE ERR FROM MASTER DATA', 'FILE NAME', '_SOURCE_ROW', 'STRUCTURE',
        ]]

        def normalize_sep(text: str) -> str:
            return re.sub(r";+", ";", re.sub(r"[^A-Za-z0-9]+", ";", str(text))).strip(";")
        
        data["PRICELIST"] = data["PRICELIST"].apply(normalize_sep)
        
        data["PRICELIST CODE"] = data["PRICELIST"].str[:4]
        
        data["PRICELIST"] = (
            data["PRICELIST"]
            .str.replace(r"(10900?|20100?|20300?|20500?)", "", regex=True)
            .str.strip()
        )
        
        self._ensure_note_err(data)
        data = self._validate_structure_gold_lv(data)
        
        def _convert_date_sp(data: Optional[pd.DataFrame]) -> Optional[pd.DataFrame]:
            data["SP START DATE"], data["SP END DATE"] = self._combine_date_columns(
                data, "SP"
            )
        
            today = pd.Timestamp.today().normalize()
        
            mask = data["SP START DATE"].notna() & (data["SP START DATE"] <= today)
            data.loc[mask, "SP START DATE"] = today + pd.Timedelta(days=1)
        
            mask_err = (
                data["SP START DATE"].notna()
                & data["SP END DATE"].notna()
                & (data["SP START DATE"] > data["SP END DATE"])
            )
        
            self._append_note_err(
                data,
                data.index[mask_err],
                "SP START DATE phải nhỏ hơn hoặc bằng SP END DATE."
            )
        
            return data
        
        data = _convert_date_sp(data)
        
        self.src_listoff = data
        
        return self
