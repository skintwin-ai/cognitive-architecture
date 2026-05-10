from __future__ import annotations

from financial_model_config import parse_output_args, resolve_output_path
from create_assumptions_sheet import main as create_assumptions
from create_operational_costs_sheet import main as create_operational_costs
from create_cogs_sheet import main as create_cogs
from create_sga_sheet import main as create_sga
from create_revenue_sheet import main as create_revenue
from create_group_model import main as create_group_model
from create_dashboard_sheet import main as create_dashboard


def main(output_path: str | None = None):
    resolved_output = str(resolve_output_path(output_path))
    create_assumptions(resolved_output)
    create_operational_costs(resolved_output)
    create_cogs(resolved_output)
    create_sga(resolved_output)
    create_revenue(resolved_output)
    create_group_model(resolved_output)
    create_dashboard(resolved_output)
    print(f'Workbook generated at {resolved_output}')
    return resolved_output


if __name__ == '__main__':
    args = parse_output_args('Generate all SkinTwin financial model sheets into a single workbook.')
    main(args.output)
