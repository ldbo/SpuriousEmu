from lark import Lark

grammar = """
start: comment
     | id

id: IDENTIFIER
comment: REM " " COMMENT_BODY

IDENTIFIER.10: /[a-zA-Z]+/
REM: "Rem"
COMMENT_BODY: /.+/


"""

text = "Rem abcd"


class DDD:
    def process(self, stream):
        for token in stream:
            print(f"Zbob: {token.type}")
            yield token

    @property
    def always_accept(self):
        return tuple()
        # return ('COMMENT_BODY', 'IDENTIFIER', 'REM')


parser = Lark(grammar, start="start", parser="lalr", debug=True, postlex=DDD())
print(parser.parse(text).pretty())
