#!/opt/libreoffice5.2/program/python
# -*- coding: utf-8 -*-
IMPLE_NAME = "pq.Tcu"
SERVICE_NAME = "pq.Tcu"
def create(ctx, *args):    
        from inoxt import component
        return component.create(ctx, *args, imple_name=IMPLE_NAME, service_name=SERVICE_NAME)
# Registration
import unohelper
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(create, IMPLE_NAME, (SERVICE_NAME,),)
