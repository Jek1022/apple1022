# -*- coding: utf-8 -*-
# Generated by Django 1.9.10 on 2017-02-27 17:12
from __future__ import unicode_literals

import datetime
from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('branch', '0002_auto_20170202_1454'),
    ]

    operations = [
        migrations.AlterField(
            model_name='branch',
            name='modifydate',
            field=models.DateTimeField(default=datetime.datetime(2017, 2, 27, 17, 12, 53, 655674)),
        ),
    ]
