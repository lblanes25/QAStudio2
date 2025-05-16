import numpy as np

class DummyParser:
    def __init__(self):
        self.complex_functions = {
            'AND': self._process_and_function,
            'IF': self._process_if_function,
        }
        self.simple_function_map = {
            'ISBLANK': 'pd.isna'
        }

    def _process_token(self, token, fields_used):
        if token not in fields_used:
            fields_used.append(token)
        return f"df['{token}']", [token]

    def _process_simple_function(self, func_name, tokens, i, fields_used):
        arg_token = tokens[i + 2]  # e.g., ISBLANK(C2)
        expr, field_list = self._process_token(arg_token, fields_used)
        return f"pd.isna({expr})", i + 4, field_list

    def _process_and_function(self, tokens, start_idx, fields_used):
        args, end_idx, new_fields = self._extract_function_args(tokens, start_idx, fields_used)
        conditions = [f"({arg})" for arg in args]
        return f"({' & '.join(conditions)})", end_idx, new_fields

    def _process_if_function(self, tokens, start_idx, fields_used):
        args, end_idx, new_fields = self._extract_function_args(tokens, start_idx, fields_used)
        if len(args) == 2:
            return f"np.where({args[0]}, {args[1]}, None)", end_idx, new_fields
        elif len(args) == 3:
            return f"np.where({args[0]}, {args[1]}, {args[2]})", end_idx, new_fields
        else:
            return "None", end_idx, new_fields

    def _extract_function_args(self, tokens, start_idx, fields_used):
        args = []
        new_fields = []
        i = start_idx
        while i < len(tokens) and tokens[i] != '(':
            i += 1
        i += 1
        paren_level = 0
        arg_tokens = []

        while i < len(tokens):
            token = tokens[i]
            if token == '(':
                paren_level += 1
                arg_tokens.append(token)
            elif token == ')':
                if paren_level == 0:
                    if arg_tokens:
                        expr = self._process_function_arg(arg_tokens, fields_used, new_fields)
                        args.append(expr)
                    i += 1
                    break
                else:
                    paren_level -= 1
                    arg_tokens.append(token)
            elif token == ',' and paren_level == 0:
                if arg_tokens:
                    expr = self._process_function_arg(arg_tokens, fields_used, new_fields)
                    args.append(expr)
                arg_tokens = []
            else:
                arg_tokens.append(token)
            i += 1

        return args, i, new_fields

    def _process_function_arg(self, tokens, fields_used, new_fields):
        if not tokens:
            return ""
        i = 0
        processed_tokens = []

        while i < len(tokens):
            token = tokens[i]
            if (i + 1 < len(tokens) and tokens[i+1] == '(' and
                (token.upper() in self.complex_functions or token.upper() in self.simple_function_map)):
                func_name = token.upper()
                if func_name in self.complex_functions:
                    handler = self.complex_functions[func_name]
                    func_expr, new_i, func_fields = handler(tokens, i, fields_used)
                    processed_tokens.append(func_expr)
                    new_fields.extend(func_fields)
                    i = new_i
                    continue
                elif func_name in self.simple_function_map:
                    func_expr, new_i, func_fields = self._process_simple_function(func_name, tokens, i, fields_used)
                    processed_tokens.append(func_expr)
                    new_fields.extend(func_fields)
                    i = new_i
                    continue
            else:
                expr, field_list = self._process_token(token, fields_used)
                processed_tokens.append(expr)
                new_fields.extend(field_list)
            i += 1

        return " ".join(processed_tokens)


def test_if_with_nested_and_and_isblank():
    parser = DummyParser()

    tokens = [
        'IF', '(', 
            'AND', '(', 
                'ISBLANK', '(', 'C2', ')', ',', 
                'C3', '=', '"yes"', 
            ')', ',', 
            '"Flag"', ',', 
            '"OK"', 
        ')'
    ]
    fields_used = []

    result_expr, end_idx, new_fields = parser._process_if_function(tokens, 0, fields_used)

    print("Parsed expression:", result_expr)
    print("New fields:", new_fields)

    assert 'pd.isna(df[\'C2\'])' in result_expr
    assert '(df[\'C3\'] == "yes")' in result_expr
    assert '"Flag"' in result_expr
    assert '"OK"' in result_expr
    assert 'C2' in new_fields
    assert 'C3' in new_fields

if __name__ == "__main__":
    test_if_with_nested_and_and_isblank()
    print("Test passed ✅")
