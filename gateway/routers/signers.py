"""GET /signers — кто подписывает документы этой компании.

Doc-V подписи не присылает: состав знает шлюз. Действие «HTTP-запрос»
может спросить его и разложить ответ по полям — например, чтобы
напечатать карточку договора или показать состав в карточке реестра.
"""
from fastapi import APIRouter, Request

router = APIRouter()


@router.get("/signers")
def resolve(request: Request, company: str = "", object: str = "",
            expense_type: str = ""):
    store = request.app.state.signers
    if not company:
        return {"companies": [
            {"company": b["company"], "object": b["object_name"], "set": b["set_name"]}
            for b in store.rules()]}
    return store.resolve(company, object, expense_type)
