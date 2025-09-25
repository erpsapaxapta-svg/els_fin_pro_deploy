from typing import List, Optional, Tuple

# TODO: replace this mock with your actual data source (DB/files/report generator)
def fetch_balance_sheet(period: str, entities: Optional[List[str]], offset: int, limit: int) -> Tuple[list, int]:
    sample = [
        {"ref": "1001", "description": "Cash and equivalents", "entity": "HO", "period": period, "amount": 12345.67},
        {"ref": "1002", "description": "Accounts receivable", "entity": "Training", "period": period, "amount": 8901.23},
    ]
    # filter by entities if provided
    if entities:
        sample = [r for r in sample if r["entity"] in set(entities)]
    total_count = len(sample)
    return sample[offset: offset + limit], total_count

