import pandas as pd
from typing import List, Dict
from collections import defaultdict


def parse_strategy_logic(df_logic: pd.DataFrame) -> Dict[str, List[str]]:
    """
    Parse the strategy logic table into a rule map for evaluation.
    Returns a dict mapping rule keys to lists of condition expressions.
    """
    rule_map = defaultdict(list)
    grouped = df_logic.groupby(["Rule Type", "Action at"], sort=False)

    for (rule_type, action_at), group_df in grouped:
        expr_parts = []

        for idx, row in group_df.iterrows():
            col_a = str(row["Column A"]).strip()
            op = str(row["Operator"]).strip()
            col_b = str(row["Column B / Value"]).strip()
            logic = str(row["Logic Type"]).strip().upper() if pd.notna(row["Logic Type"]) else ""

            # Determine if col_b is a value or a column
            if col_b.replace(".", "", 1).isdigit():
                cond = f"(row['{col_a}'] {op} {col_b})"
            else:
                cond = f"(row['{col_a}'] {op} row['{col_b}'])"

            expr_parts.append(cond)

            # Add logical operator if not END
            if logic in {"AND", "OR"}:
                expr_parts.append(logic.lower())
            elif logic not in {"", "END"}:
                print(f"⚠️ Unknown logic type: '{logic}' in row {idx}")

        # Remove trailing logic operator if present
        if expr_parts and expr_parts[-1] in {"and", "or"}:
            expr_parts = expr_parts[:-1]

        key = f"{rule_type.strip()}_{action_at.strip()}"
        rule_map[key].append(" ".join(expr_parts))

    return rule_map


def evaluate_conditions(row: pd.Series, conditions: List[str]) -> bool:
    """
    Evaluate if a set of conditions is satisfied for a given row.
    """
    try:
        expr = " ".join(conditions)
        # Ensure all row[...] are scalars, not Series
        # If any value is a Series, use .item() if length 1, else warn and skip
        class SafeRow(dict):
            def __getitem__(self, key):
                val = super().__getitem__(key)
                if isinstance(val, pd.Series):
                    if len(val) == 1:
                        return val.item()
                    else:
                        print(f"[Strategy] Ambiguous value for '{key}' in row: Series of length {len(val)}. Skipping.")
                        raise ValueError(f"Ambiguous value for '{key}' in row.")
                return val
        safe_row = SafeRow(row)
        return eval(expr, {"row": safe_row})
    except Exception as e:
        print(f"❌ Evaluation error: {e}, expr: {expr}")
        return False


def strategy_from_logic(df: pd.DataFrame, rules: Dict[str, List[str]]) -> pd.DataFrame:
    """
    Apply strategy logic to a DataFrame and return a DataFrame of trades with PnL and triggers.
    """
    results = []
    position = None  # None or dict with entry info
    i = 0
    while i < len(df):
        row = df.iloc[i]
        date = row["Date"]

        # If no open position, check entry
        if position is None:
            action = None
            entry_field = None
            entry_price = None
            entry_triggered = False
            for rule_key in rules:
                if rule_key.startswith("Enter-"):
                    rule_type, entry_field_candidate = rule_key.split("_")
                    try:
                        triggered = evaluate_conditions(row, rules[rule_key])
                    except Exception:
                        continue
                    if triggered:
                        if "Buy" in rule_type:
                            action = "Buy"
                        elif "Sell" in rule_type:
                            action = "Sell"
                        entry_field = entry_field_candidate
                        entry_price = row.get(entry_field, row["Open"])
                        position = {
                            "Action": action,
                            "Entry": entry_price,
                            "EntryDate": date,
                            "EntryIdx": i,
                            "EntryField": entry_field,
                            "Stop": False,
                            "TakeProfit": False
                        }
                        entry_triggered = True
                        break
            # After entry, immediately check exit/stop/take profit in the same row
            if entry_triggered:
                exit_price = None
                stop = False
                take_profit = False
                # Step 1: Stop loss
                stop_prefix = "StopLoss-long" if action == "Buy" else "StopLoss-short"
                for rule_key in rules:
                    if rule_key.startswith(stop_prefix):
                        try:
                            stop_triggered = evaluate_conditions(row, rules[rule_key])
                            if stop_triggered:
                                _, action_at = rule_key.split("_", 1)
                                exit_price = row.get(action_at, row["Close"])
                                stop = True
                                break
                        except Exception:
                            pass
                # Step 2: Take profit
                if not stop:
                    tp_prefix = "TakeProfit-long" if action == "Buy" else "TakeProfit-short"
                    for rule_key in rules:
                        if rule_key.startswith(tp_prefix):
                            try:
                                tp_triggered = evaluate_conditions(row, rules[rule_key])
                                if tp_triggered:
                                    _, action_at = rule_key.split("_", 1)
                                    exit_price = row.get(action_at, row["Close"])
                                    take_profit = True
                                    break
                            except Exception:
                                pass
                # Step 3: Exit
                if not stop and not take_profit:
                    exit_prefix = "Exit-long" if action == "Buy" else "Exit-short"
                    for rule_key in rules:
                        if rule_key.startswith(exit_prefix):
                            try:
                                exit_triggered = evaluate_conditions(row, rules[rule_key])
                                if exit_triggered:
                                    _, action_at = rule_key.split("_", 1)
                                    exit_price = row.get(action_at, row["Close"])
                                    break
                            except Exception:
                                pass
                if exit_price is not None:
                    pnl = (exit_price - entry_price) if action == "Buy" else (entry_price - exit_price)
                    results.append({
                        "EntryDate": date,
                        "ExitDate": date,
                        "Action": action,
                        "Entry": entry_price,
                        "Exit": exit_price,
                        "Stop Triggered": stop,
                        "Take Profit Triggered": take_profit,
                        "PnL": pnl
                    })
                    position = None
                    i += 1
                    continue
            i += 1
            continue

        # If position is open, check exit/stop/take profit in current row
        exit_price = None
        stop = False
        take_profit = False
        action = position["Action"]
        entry_price = position["Entry"]
        entry_date = position["EntryDate"]

        # Step 1: Stop loss
        stop_prefix = "StopLoss-long" if action == "Buy" else "StopLoss-short"
        for rule_key in rules:
            if rule_key.startswith(stop_prefix):
                try:
                    stop_triggered = evaluate_conditions(row, rules[rule_key])
                    if stop_triggered:
                        _, action_at = rule_key.split("_", 1)
                        exit_price = row.get(action_at, row["Close"])
                        stop = True
                        break
                except Exception:
                    pass

        # Step 2: Take profit
        if not stop:
            tp_prefix = "TakeProfit-long" if action == "Buy" else "TakeProfit-short"
            for rule_key in rules:
                if rule_key.startswith(tp_prefix):
                    try:
                        tp_triggered = evaluate_conditions(row, rules[rule_key])
                        if tp_triggered:
                            _, action_at = rule_key.split("_", 1)
                            exit_price = row.get(action_at, row["Close"])
                            take_profit = True
                            break
                    except Exception:
                        pass

        # Step 3: Exit
        if not stop and not take_profit:
            exit_prefix = "Exit-long" if action == "Buy" else "Exit-short"
            for rule_key in rules:
                if rule_key.startswith(exit_prefix):
                    try:
                        exit_triggered = evaluate_conditions(row, rules[rule_key])
                        if exit_triggered:
                            _, action_at = rule_key.split("_", 1)
                            exit_price = row.get(action_at, row["Close"])
                            break
                    except Exception:
                        pass

        # Record trade only if exit_price is not None
        if exit_price is not None:
            pnl = (exit_price - entry_price) if action == "Buy" else (entry_price - exit_price)
            results.append({
                "EntryDate": entry_date,
                "ExitDate": date,
                "Action": action,
                "Entry": entry_price,
                "Exit": exit_price,
                "Stop Triggered": stop,
                "Take Profit Triggered": take_profit,
                "PnL": pnl
            })
            position = None  # Reset position
        i += 1
    return pd.DataFrame(results)