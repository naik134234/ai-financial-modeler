"""
QA Validator Agent
Validates financial model for accuracy, consistency, and reasonableness
"""

from typing import Dict, Any, List, Tuple, Optional
import logging

logger = logging.getLogger(__name__)


class ValidationError:
    """Represents a validation error or warning"""
    
    def __init__(
        self,
        severity: str,  # 'error', 'warning', 'info'
        category: str,  # 'balance', 'formula', 'ratio', 'assumption'
        message: str,
        location: Optional[str] = None,
        value: Optional[Any] = None,
        expected: Optional[Any] = None
    ):
        self.severity = severity
        self.category = category
        self.message = message
        self.location = location
        self.value = value
        self.expected = expected
    
    def to_dict(self) -> Dict[str, Any]:
        return {
            'severity': self.severity,
            'category': self.category,
            'message': self.message,
            'location': self.location,
            'value': self.value,
            'expected': self.expected,
        }


class QAValidator:
    """Validates financial models for accuracy and consistency"""
    
    # Tolerance for balance checks (0.01 = 1%)
    BALANCE_TOLERANCE = 0.01
    
    # Reasonable ratio ranges for sanity checks
    RATIO_RANGES = {
        'gross_margin': (0.0, 1.0),
        'ebitda_margin': (-0.5, 0.8),
        'net_margin': (-1.0, 0.5),
        'roe': (-1.0, 1.0),
        'roa': (-0.5, 0.5),
        'current_ratio': (0.1, 10.0),
        'debt_to_equity': (0.0, 20.0),
        'revenue_growth': (-0.5, 2.0),
    }
    
    # Power sector specific ranges
    POWER_RATIO_RANGES = {
        'plf': (0.3, 0.95),
        'dscr': (0.8, 5.0),
        'llcr': (0.8, 5.0),
    }
    
    def __init__(self, industry_code: str = 'general'):
        """Initialize validator with industry context"""
        self.industry_code = industry_code
        self.errors: List[ValidationError] = []
    
    def validate_model(self, model_data: Dict[str, Any]) -> Tuple[bool, List[Dict[str, Any]]]:
        """
        Validate the entire financial model
        
        Args:
            model_data: Dictionary containing all model data
        
        Returns:
            Tuple of (is_valid, list of validation errors/warnings)
        """
        self.errors = []
        
        # Run all validation checks
        self._validate_balance_sheet(model_data.get('balance_sheet', {}))
        self._validate_cash_flow(model_data.get('cash_flow', {}), model_data.get('balance_sheet', {}))
        self._validate_income_statement(model_data.get('income_statement', {}))
        self._validate_ratios(model_data.get('ratios', {}))
        self._validate_assumptions(model_data.get('assumptions', {}))
        
        if self.industry_code == 'power':
            self._validate_power_specific(model_data)
        
        # Check for critical errors
        has_errors = any(e.severity == 'error' for e in self.errors)
        
        return (not has_errors, [e.to_dict() for e in self.errors])
    
    def _validate_balance_sheet(self, bs_data: Dict[str, Any]) -> None:
        """Validate balance sheet equation: Assets = Liabilities + Equity"""
        if not bs_data:
            return
        
        # For each period (historical and forecast)
        for period, data in bs_data.items():
            if not isinstance(data, dict):
                continue
            
            assets = data.get('total_assets', 0) or 0
            liabilities = data.get('total_liabilities', 0) or 0
            equity = data.get('total_equity', 0) or 0
            
            if assets == 0:
                continue
            
            # Check balance
            diff = abs(assets - (liabilities + equity))
            tolerance = abs(assets) * self.BALANCE_TOLERANCE
            
            if diff > tolerance:
                self.errors.append(ValidationError(
                    severity='error',
                    category='balance',
                    message=f'Balance sheet does not balance',
                    location=f'Balance Sheet - {period}',
                    value={'assets': assets, 'liabilities': liabilities, 'equity': equity},
                    expected=f'Assets ({assets:,.0f}) = Liabilities ({liabilities:,.0f}) + Equity ({equity:,.0f})'
                ))
            
            # Check for negative equity (warning)
            if equity < 0:
                self.errors.append(ValidationError(
                    severity='warning',
                    category='balance',
                    message='Negative equity detected',
                    location=f'Balance Sheet - {period}',
                    value=equity
                ))
    
    def _validate_cash_flow(self, cf_data: Dict[str, Any], bs_data: Dict[str, Any]) -> None:
        """Validate cash flow reconciliation"""
        if not cf_data or not bs_data:
            return
        
        periods = sorted([p for p in cf_data.keys() if isinstance(cf_data[p], dict)])
        
        for i, period in enumerate(periods):
            cf = cf_data[period]
            
            # Check OCF + ICF + FCF = Net Cash Flow
            ocf = cf.get('operating_cash_flow', 0) or 0
            icf = cf.get('investing_cash_flow', 0) or 0
            fcf = cf.get('financing_cash_flow', 0) or 0
            net_cf = cf.get('net_cash_flow', 0) or 0
            
            calculated_net = ocf + icf + fcf
            if net_cf != 0 and abs(calculated_net - net_cf) > abs(net_cf) * self.BALANCE_TOLERANCE:
                self.errors.append(ValidationError(
                    severity='error',
                    category='balance',
                    message='Cash flow components do not reconcile',
                    location=f'Cash Flow - {period}',
                    value={'OCF': ocf, 'ICF': icf, 'FCF': fcf, 'Net': net_cf},
                    expected=f'OCF + ICF + FCF should equal Net CF'
                ))
            
            # Check cash reconciliation with balance sheet
            if period in bs_data and isinstance(bs_data[period], dict):
                closing_cash = bs_data[period].get('cash', 0) or 0
                
                if i > 0 and periods[i-1] in bs_data:
                    opening_cash = bs_data[periods[i-1]].get('cash', 0) or 0
                    expected_closing = opening_cash + net_cf
                    
                    if closing_cash != 0 and abs(expected_closing - closing_cash) > abs(closing_cash) * self.BALANCE_TOLERANCE:
                        self.errors.append(ValidationError(
                            severity='warning',
                            category='balance',
                            message='Cash balance does not reconcile with cash flow',
                            location=f'Cash Flow - {period}',
                            value={'opening': opening_cash, 'net_cf': net_cf, 'closing': closing_cash},
                            expected=f'Opening + Net CF = Closing'
                        ))
    
    def _validate_income_statement(self, is_data: Dict[str, Any]) -> None:
        """Validate income statement logic"""
        if not is_data:
            return
        
        for period, data in is_data.items():
            if not isinstance(data, dict):
                continue
            
            revenue = data.get('revenue', 0) or 0
            gross_profit = data.get('gross_profit', 0) or 0
            ebitda = data.get('ebitda', 0) or 0
            net_income = data.get('net_income', 0) or 0
            
            if revenue <= 0:
                continue
            
            # Check gross profit <= revenue
            if gross_profit > revenue * 1.01:  # 1% tolerance
                self.errors.append(ValidationError(
                    severity='error',
                    category='formula',
                    message='Gross profit exceeds revenue',
                    location=f'Income Statement - {period}',
                    value={'revenue': revenue, 'gross_profit': gross_profit}
                ))
            
            # Check EBITDA <= gross profit (generally)
            if ebitda > gross_profit * 1.1:  # 10% tolerance for other income
                self.errors.append(ValidationError(
                    severity='warning',
                    category='formula',
                    message='EBITDA exceeds gross profit significantly',
                    location=f'Income Statement - {period}',
                    value={'gross_profit': gross_profit, 'ebitda': ebitda}
                ))
    
    def _validate_ratios(self, ratios: Dict[str, Any]) -> None:
        """Validate financial ratios are within reasonable ranges"""
        for ratio_name, (min_val, max_val) in self.RATIO_RANGES.items():
            if ratio_name in ratios:
                value = ratios[ratio_name]
                if value is not None and (value < min_val or value > max_val):
                    self.errors.append(ValidationError(
                        severity='warning',
                        category='ratio',
                        message=f'{ratio_name} is outside normal range',
                        location='Ratios',
                        value=value,
                        expected=f'Expected between {min_val} and {max_val}'
                    ))
    
    def _validate_assumptions(self, assumptions: Dict[str, Any]) -> None:
        """Validate model assumptions are reasonable"""
        # Check revenue growth
        growth = assumptions.get('revenue_growth_rate')
        if growth is not None:
            if growth > 0.5:  # 50%
                self.errors.append(ValidationError(
                    severity='warning',
                    category='assumption',
                    message='Revenue growth rate seems aggressive',
                    location='Assumptions',
                    value=growth,
                    expected='Typically < 50% for mature companies'
                ))
            if growth < -0.3:  # -30%
                self.errors.append(ValidationError(
                    severity='warning',
                    category='assumption',
                    message='Revenue decline seems severe',
                    location='Assumptions',
                    value=growth
                ))
        
        # Check terminal growth
        terminal_growth = assumptions.get('terminal_growth')
        if terminal_growth is not None:
            if terminal_growth > 0.06:  # 6%
                self.errors.append(ValidationError(
                    severity='warning',
                    category='assumption',
                    message='Terminal growth rate exceeds long-term GDP growth',
                    location='Assumptions - Valuation',
                    value=terminal_growth,
                    expected='Should be <= long-term nominal GDP growth (4-6%)'
                ))
        
        # Check WACC
        wacc = assumptions.get('wacc')
        if wacc is not None:
            if wacc < 0.05 or wacc > 0.25:
                self.errors.append(ValidationError(
                    severity='warning',
                    category='assumption',
                    message='WACC seems unusual',
                    location='Assumptions - Valuation',
                    value=wacc,
                    expected='Typically between 5% and 25%'
                ))
    
    def _validate_power_specific(self, model_data: Dict[str, Any]) -> None:
        """Validate power sector specific metrics"""
        power_data = model_data.get('power_operations', {})
        
        # Check PLF
        plf = power_data.get('plf')
        if plf is not None:
            min_plf, max_plf = self.POWER_RATIO_RANGES['plf']
            if plf < min_plf or plf > max_plf:
                self.errors.append(ValidationError(
                    severity='warning',
                    category='ratio',
                    message='PLF is outside typical range',
                    location='Power Operations',
                    value=plf,
                    expected=f'Expected between {min_plf*100}% and {max_plf*100}%'
                ))
        
        # Check DSCR
        dscr = power_data.get('dscr')
        if dscr is not None:
            if dscr < 1.0:
                self.errors.append(ValidationError(
                    severity='error',
                    category='ratio',
                    message='DSCR below 1.0 indicates insufficient cash flow for debt service',
                    location='Power Operations',
                    value=dscr,
                    expected='DSCR should be >= 1.0 for debt serviceability'
                ))
            elif dscr < 1.2:
                self.errors.append(ValidationError(
                    severity='warning',
                    category='ratio',
                    message='DSCR is low, may trigger covenant issues',
                    location='Power Operations',
                    value=dscr,
                    expected='DSCR typically required >= 1.2 by lenders'
                ))


def validate_financial_model(
    model_data: Dict[str, Any],
    industry_code: str = 'general'
) -> Tuple[bool, List[Dict[str, Any]]]:
    """
    Convenience function to validate a financial model
    
    Args:
        model_data: Dictionary containing all model data
        industry_code: Industry code for sector-specific validation
    
    Returns:
        Tuple of (is_valid, list of validation errors/warnings)
    """
    validator = QAValidator(industry_code)
    return validator.validate_model(model_data)


if __name__ == "__main__":
    # Test validation
    test_model = {
        'balance_sheet': {
            'FY2024': {
                'total_assets': 100000,
                'total_liabilities': 60000,
                'total_equity': 40000,  # Should balance
                'cash': 5000,
            }
        },
        'income_statement': {
            'FY2024': {
                'revenue': 50000,
                'gross_profit': 20000,
                'ebitda': 15000,
                'net_income': 8000,
            }
        },
        'assumptions': {
            'revenue_growth_rate': 0.15,
            'terminal_growth': 0.04,
            'wacc': 0.12,
        }
    }
    
    is_valid, errors = validate_financial_model(test_model)
    print(f"Model valid: {is_valid}")
    print(f"Errors/Warnings: {len(errors)}")
    for error in errors:
        print(f"  [{error['severity']}] {error['message']}")
