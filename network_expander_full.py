"""Standalone network parsing, expansion and Stage 1 network validation.

Dependencies: pandas, openpyxl (for from_excel).
Extracted from the current Template_ETL; does not import project modules.
Site Group assignment and non-network business validation are not included.
"""
from collections import Counter
import re
from typing import List, Optional
import pandas as pd

required_network = ['SITE', 'NATIONAL_SITE', 'GROUP_SITE', 'REGION_SITE', 'ACTIVE', 'DISCOUNT']

class NetworkExpander:
    GROUP_COLS = ["FILE NAME", "GOLD CODE", "LV", "LU"]

    def __init__(self, dict_network: dict):
        self.dict_network = dict_network

    @classmethod
    def from_excel(cls, path):
        """Read network-configure and build all eight dictionary keys."""
        instance = cls({})
        instance.path_plan = path
        instance._load_network()
        return instance

    def expand(self, expression, deduplicate=True, discount=False):
        """Normalize and expand a single expression; does not run row validation."""
        expression = self._normalize_network_punctuation(expression)
        if discount:
            return self._extract_discount_network(expression)
        return self._extract_network(expression, deduplicate=deduplicate)

    def validate(self, data: pd.DataFrame) -> pd.DataFrame:
        """Return a copy with normalized inputs, expanded columns and error notes.

        Required columns: FILE NAME, GOLD CODE, LV, LU,
        PURCHASE NETWORK, GOLD PROMO NETWORK.
        This runs network validation only, not the entire Stage 1 pipeline.
        """
        self._check_required_columns(data, [*self.GROUP_COLS, "PURCHASE NETWORK", "GOLD PROMO NETWORK"])
        if not data.index.is_unique:
            raise ValueError("DataFrame index must be unique; use reset_index(drop=True).")
        return self._ppNetwork_gpNetwork(self._check_network(data.copy()))

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

    def _unique_sorted_sites(self, value) -> tuple:
        return tuple(sorted(set(self._parse_sites(value)), key=self._sort_key))

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

