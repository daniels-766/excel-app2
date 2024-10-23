from celery import Celery

# Set up the Celery instance
def make_celery(app):
    celery = Celery(app.import_name, backend='redis://localhost:6379/0', broker='redis://localhost:6379/0')
    celery.conf.update(app.config)
    return celery
