"""Swagger/OpenAPI документация для REST API."""

from flask import Blueprint, jsonify, render_template_string

api_docs_bp = Blueprint('api_docs', __name__, url_prefix='/api/docs')


SWAGGER_SPEC = {
    'openapi': '3.0.0',
    'info': {
        'title': 'KP Generator API',
        'version': '1.0.0',
        'description': 'REST API для генератора коммерческих предложений',
    },
    'servers': [
        {
            'url': '/api/v1',
            'description': 'Production server'
        }
    ],
    'paths': {
        '/health': {
            'get': {
                'summary': 'Проверка работоспособности API',
                'tags': ['Health'],
                'responses': {
                    '200': {
                        'description': 'API работает',
                        'content': {
                            'application/json': {
                                'schema': {
                                    'type': 'object',
                                    'properties': {
                                        'status': {'type': 'string'},
                                        'version': {'type': 'string'},
                                        'service': {'type': 'string'}
                                    }
                                }
                            }
                        }
                    }
                }
            }
        },
        '/generations': {
            'get': {
                'summary': 'Получить список генераций',
                'tags': ['Generations'],
                'security': [{'BearerAuth': []}],
                'parameters': [
                    {
                        'name': 'page',
                        'in': 'query',
                        'schema': {'type': 'integer', 'default': 1}
                    },
                    {
                        'name': 'per_page',
                        'in': 'query',
                        'schema': {'type': 'integer', 'default': 25}
                    },
                    {
                        'name': 'date_from',
                        'in': 'query',
                        'schema': {'type': 'string', 'format': 'date'}
                    },
                    {
                        'name': 'date_to',
                        'in': 'query',
                        'schema': {'type': 'string', 'format': 'date'}
                    }
                ],
                'responses': {
                    '200': {
                        'description': 'Список генераций',
                        'content': {
                            'application/json': {
                                'schema': {
                                    'type': 'object',
                                    'properties': {
                                        'items': {'type': 'array'},
                                        'pagination': {'type': 'object'}
                                    }
                                }
                            }
                        }
                    }
                }
            }
        },
        '/generations/{generation_id}': {
            'get': {
                'summary': 'Получить детальную информацию о генерации',
                'tags': ['Generations'],
                'security': [{'BearerAuth': []}],
                'parameters': [
                    {
                        'name': 'generation_id',
                        'in': 'path',
                        'required': True,
                        'schema': {'type': 'integer'}
                    }
                ],
                'responses': {
                    '200': {
                        'description': 'Детальная информация о генерации'
                    },
                    '404': {
                        'description': 'Генерация не найдена'
                    }
                }
            }
        },
        '/calculate': {
            'post': {
                'summary': 'Рассчитать цены для позиций',
                'tags': ['Calculations'],
                'security': [{'BearerAuth': []}],
                'requestBody': {
                    'required': True,
                    'content': {
                        'application/json': {
                            'schema': {
                                'type': 'object',
                                'required': ['positions', 'logistics_rub', 'delivery_time', 'margin_percent'],
                                'properties': {
                                    'positions': {
                                        'type': 'array',
                                        'items': {'type': 'object'}
                                    },
                                    'logistics_rub': {'type': 'number'},
                                    'delivery_time': {'type': 'integer'},
                                    'margin_percent': {'type': 'number'}
                                }
                            }
                        }
                    }
                },
                'responses': {
                    '200': {
                        'description': 'Результаты расчета'
                    },
                    '400': {
                        'description': 'Ошибка валидации'
                    }
                }
            }
        },
        '/logistics/calculate': {
            'post': {
                'summary': 'Рассчитать стоимость логистики',
                'tags': ['Logistics'],
                'requestBody': {
                    'required': True,
                    'content': {
                        'application/json': {
                            'schema': {
                                'type': 'object',
                                'required': ['weight_kg', 'city_price'],
                                'properties': {
                                    'weight_kg': {'type': 'number'},
                                    'city_price': {'type': 'number'},
                                    'transport_type': {'type': 'string', 'enum': ['truck', 'trail']},
                                    'city_name': {'type': 'string'},
                                    'is_main_route': {'type': 'boolean'},
                                    'distance_from_ekb_km': {'type': 'integer'},
                                    'length_mm': {'type': 'number'},
                                    'width_mm': {'type': 'number'},
                                    'height_mm': {'type': 'number'}
                                }
                            }
                        }
                    }
                },
                'responses': {
                    '200': {
                        'description': 'Результаты расчета логистики'
                    }
                }
            }
        },
        '/duty/search': {
            'get': {
                'summary': 'Поиск ставок пошлин',
                'tags': ['Reference Data'],
                'parameters': [
                    {
                        'name': 'query',
                        'in': 'query',
                        'schema': {'type': 'string'}
                    },
                    {
                        'name': 'limit',
                        'in': 'query',
                        'schema': {'type': 'integer', 'default': 50}
                    }
                ],
                'responses': {
                    '200': {
                        'description': 'Список найденных ставок пошлин'
                    }
                }
            }
        },
        '/materials/gb': {
            'get': {
                'summary': 'Получить список аналогов материалов GB',
                'tags': ['Reference Data'],
                'parameters': [
                    {
                        'name': 'search',
                        'in': 'query',
                        'schema': {'type': 'string'}
                    }
                ],
                'responses': {
                    '200': {
                        'description': 'Список материалов'
                    }
                }
            }
        },
        '/logistics/cities': {
            'get': {
                'summary': 'Получить список городов логистики',
                'tags': ['Reference Data'],
                'responses': {
                    '200': {
                        'description': 'Список городов'
                    }
                }
            }
        }
    },
    'components': {
        'securitySchemes': {
            'BearerAuth': {
                'type': 'http',
                'scheme': 'bearer',
                'bearerFormat': 'JWT',
                'description': 'Требуется авторизация через Flask-Login'
            }
        }
    },
    'tags': [
        {'name': 'Health', 'description': 'Проверка работоспособности'},
        {'name': 'Generations', 'description': 'Управление генерациями КП'},
        {'name': 'Calculations', 'description': 'Расчеты цен и маржи'},
        {'name': 'Logistics', 'description': 'Расчет логистики'},
        {'name': 'Reference Data', 'description': 'Справочные данные'},
    ]
}


@api_docs_bp.route('/swagger.json')
def swagger_json():
    """Возвращает Swagger спецификацию в формате JSON."""
    return jsonify(SWAGGER_SPEC)


@api_docs_bp.route('/')
def swagger_ui():
    """Отображает Swagger UI для интерактивной документации."""
    swagger_ui_html = """
    <!DOCTYPE html>
    <html>
    <head>
        <title>KP Generator API Documentation</title>
        <link rel="stylesheet" type="text/css" href="https://unpkg.com/swagger-ui-dist@5.0.0/swagger-ui.css" />
        <style>
            html { box-sizing: border-box; overflow: -moz-scrollbars-vertical; overflow-y: scroll; }
            *, *:before, *:after { box-sizing: inherit; }
            body { margin:0; background: #fafafa; }
        </style>
    </head>
    <body>
        <div id="swagger-ui"></div>
        <script src="https://unpkg.com/swagger-ui-dist@5.0.0/swagger-ui-bundle.js"></script>
        <script src="https://unpkg.com/swagger-ui-dist@5.0.0/swagger-ui-standalone-preset.js"></script>
        <script>
            window.onload = function() {
                const ui = SwaggerUIBundle({
                    url: "/api/docs/swagger.json",
                    dom_id: '#swagger-ui',
                    deepLinking: true,
                    presets: [
                        SwaggerUIBundle.presets.apis,
                        SwaggerUIStandalonePreset
                    ],
                    plugins: [
                        SwaggerUIBundle.plugins.DownloadUrl
                    ],
                    layout: "StandaloneLayout"
                });
            };
        </script>
    </body>
    </html>
    """
    return render_template_string(swagger_ui_html)

