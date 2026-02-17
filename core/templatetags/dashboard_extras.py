from django import template         # type: ignore

register = template.Library()

@register.simple_tag(takes_context=True)
def update_params(context, **kwargs):
    """
        Returns the current URL query string with updated parameters.
        Usage: <a href="?{% update_params page=2 %}">Next</a>
        Keeps existing filters (like start_date, sbu) while changing one specific param.
    """
    query = context['request'].GET.copy()
    for key, value in kwargs.items():
        query[key] = value
    return query.urlencode()

@register.filter
def get_item(dictionary, key):
    return dictionary.get(key)