# -*- coding: utf-8 -*-
# Generated by Django 1.9.10 on 2017-04-21 13:51
from __future__ import unicode_literals

import datetime
from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('unitofmeasure', '0004_auto_20170411_1434'),
    ]

    operations = [
        migrations.AlterField(
            model_name='unitofmeasure',
            name='modifydate',
            field=models.DateTimeField(default=datetime.datetime(2017, 4, 21, 13, 51, 25, 760000)),
        ),
    ]
