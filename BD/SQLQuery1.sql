select id_colilla_empleado, Fecha_inicial, Fecha_final, (Nombre + ' ' + Apellido) 
                                    nombre, Numero_documento, Salario_base, DATEDIFF(day, Fecha_inicial, Fecha_final) dias, Basico, Aux_transporte, 
				                    Total_Hrs, Total_Rec, Incapacidades, Licencias, Otros_devengados, Total_devengado, Salud, Pension, Otras_deducciones,
                                    Retencion, Prestamos, Total_deducciones, Neto_pagar from Colilla_pago where Fecha_inicial >=
                                    '2023-09-30' and Fecha_final <= '2023-10-15'

select * from Colilla_pago
   