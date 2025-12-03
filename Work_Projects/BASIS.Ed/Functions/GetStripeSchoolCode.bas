Public Function GetStripeSchoolCode(ByVal DisbursementID As Variant) As String

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''-------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''' Purpose and Updates Log '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''-------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' =============================================================
    '                            PURPOSE
    ' =============================================================
        ' This function assigns the correct School Code based on the Stripe 'Disbursement ID'.
        ' It is used when Click & Pledge reports do not include the School Code for a particular disbursement.

    ' =============================================================
    '                             NOTES
    ' =============================================================
        ' This mapping is intentionally hard-coded per system design.
        ' The table is maintained on a monthly basis as new Stripe 'Disbursement IDs' without corresponding School Codes appear in Click & Pledge reports.

    ' =============================================================
    '                           UPDATE LOG
    ' -------------------------------------------------------------
    '                     (LAST UPDATED: 2025.11.20)
    ' =============================================================
        ' Initial Creation Date  : 2025.11.20
        ' Production Rollout Date: 2025.11.20

        ' Updates:
            ' 2025.12.02 - Added new Disbursement IDs from November 2025 Click & Pledge sync that were missing School Codes.

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''-----------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''' Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''-----------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Declare the 'key' variable
        Dim key As String
        
    ' Normalize the incoming 'DisbursementID' value (trim + convert to string)
        key = Trim$(CStr(DisbursementID))

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''-----------------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''' Map the 'DisbursementID' to School Code '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''-----------------------------------------'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Select Case key
        ' ==============================
        '             2024.05
        ' ==============================
            Case "po_1PGwRCHBO4JggQcHIB3hqmg5": GetStripeSchoolCode = "62053"        ' (PHXN)
            
        ' ==============================
        '             2024.06
        ' ==============================
            Case "po_1PSATGHBO4JggQcHnnuPuQJj": GetStripeSchoolCode = "62053"        ' (PHXN)
            
        ' ==============================
        '             2024.07
        ' ==============================
            Case "po_1Pd4aCHBO4JggQcHbPwNRqXf": GetStripeSchoolCode = "62053"        ' (PHXN)
            
        ' ==============================
        '             2024.09
        ' ==============================
            Case "po_1PzVseHBO4JggQcHLEfypsUm": GetStripeSchoolCode = "62053"        ' (PHXN)
            
        ' ==============================
        '             2025.06
        ' ==============================
            Case "po_1RaSpAHBO4JggQcHzyiCyny0": GetStripeSchoolCode = "62053"        ' (PHXN)
            Case "po_1RaSVEQahEPZfhTMdRiEM5nM": GetStripeSchoolCode = "62375"        ' (PLN)
            Case "po_1RaSnjQkAnjuwifxEDgXO8z4": GetStripeSchoolCode = "62376"        ' (RCH)
            
        ' ==============================
        '             2025.07
        ' ==============================
            Case "po_1RlKj6HBO4JggQcHFZLtTYNh": GetStripeSchoolCode = "62053"        ' (PHXN)
            Case "po_1RlLBbQahEPZfhTM5jVMEzCv": GetStripeSchoolCode = "62375"        ' (PLN)
            Case "po_1RlL1QQkAnjuwifxSJ2rKZKH": GetStripeSchoolCode = "62376"        ' (RCH)
            
        ' ==============================
        '             2025.08
        ' ==============================
            Case "po_1Rwv2EQYaS7mQRKSfVjzn2Vh": GetStripeSchoolCode = "61096"        ' (AUSP)
            Case "po_1Rwv3vHoHDbMF2miFMmm6jcn": GetStripeSchoolCode = "35412"        ' (BCSI)
            Case "po_1RwvEqQYL6jHaUXMN6k0AJuc": GetStripeSchoolCode = "61620"        ' (BEN)
            Case "po_1RwvQKH41tYMeeSKNneHkmWR": GetStripeSchoolCode = "61367"        ' (BRMC)
            Case "po_1Rwv4FQjF0hKzNBk1m2CEdxL": GetStripeSchoolCode = "35422"        ' (CHPS)
            Case "po_1RwvDnQeaxl1oH0SwtYsHW6U": GetStripeSchoolCode = "61773"        ' (CPK)
            Case "po_1Rwv3WH0K05F9mgHGoySQvGe": GetStripeSchoolCode = "61621"        ' (PFL)
            Case "po_1RwvAtHePa3cGmcJNkanmtAd": GetStripeSchoolCode = "35431"        ' (PHXC)
            Case "po_1RwvGoHW8Eo1hm07IUi3LXhb": GetStripeSchoolCode = "38346"        ' (PHXS)
            Case "po_1RwvDXQkAnjuwifxZSRB1rFz": GetStripeSchoolCode = "62376"        ' (RCH)
            Case "po_1RwvCiHQOHalnvHrvwFOgHI3": GetStripeSchoolCode = "61097"        ' (SANE)
            Case "po_1Rwv1hHZSXdjvotcqUONSyKw": GetStripeSchoolCode = "38511"        ' (SAS)
            Case "po_1Rwv83QgoLYKu1DKk813Hk03": GetStripeSchoolCode = "61098"        ' (SPNE)
            
        ' ==============================
        '             2025.09
        ' ==============================
            Case "po_1S7nVQQhnfWS9EivIqgfmiWu": GetStripeSchoolCode = "61369"        ' (AUS)
            Case "po_1S7nP3QYaS7mQRKSS0vmnmN1": GetStripeSchoolCode = "61096"        ' (AUSP)
            Case "po_1S7nOSH9wkJRSqOam3hzGy05": GetStripeSchoolCode = "50884"        ' (BBRM)
            Case "po_1S7nHeHoHDbMF2mijdAKvCRx": GetStripeSchoolCode = "35412"        ' (BCSI)
            Case "po_1S7nPjHzt9XZ7h60nccupdTe": GetStripeSchoolCode = "35419"        ' (CHD)
            Case "po_1S7nH4HgYDI9wgzhgBvmR1vB": GetStripeSchoolCode = "35426"        ' (GDY)
            Case "po_1S7nL8H46HhwWaTC55GOOm7z": GetStripeSchoolCode = "35427"        ' (MES)
            Case "po_1S7nLFHvTxXqOAZseOMBW7Ha": GetStripeSchoolCode = "35428"        ' (OV)
            Case "po_1S7nMNHZjj5CWGhpxcoah7PO": GetStripeSchoolCode = "35439"        ' (OVP)
            Case "po_1S7nOGHbVYyrxVVwA4EwISE2": GetStripeSchoolCode = "35429"        ' (PEO)
            Case "po_1S7nO7H0K05F9mgHIbhyyjVv": GetStripeSchoolCode = "61621"        ' (PFL)
            Case "po_1S7oREHBO4JggQcHBgwcHNjz": GetStripeSchoolCode = "62053"        ' (PHXN)
            Case "po_1S7nXoHW8Eo1hm07NcjCkXX8": GetStripeSchoolCode = "38346"        ' (PHXS)
            Case "po_1S7odWQahEPZfhTMJx31Lkhu": GetStripeSchoolCode = "62375"        ' (PLN)
            Case "po_1S7nK7Hkv0Z5SiNeuMFfF2uG": GetStripeSchoolCode = "35432"        ' (PRE)
            Case "po_1S7oIiQkAnjuwifxZyz7GuNF": GetStripeSchoolCode = "62376"        ' (RCH)
            Case "po_1S7nUGHC1Kl6KxgMp32HodWA": GetStripeSchoolCode = "35441"        ' (SAMC)
            Case "po_1S7nMrHQOHalnvHrH52Rf2ys": GetStripeSchoolCode = "61097"        ' (SANE)
            Case "po_1S7nPDHZSXdjvotc6vEnF60A": GetStripeSchoolCode = "38511"        ' (SAS)
            Case "po_1S7nYBQgoLYKu1DKmazQKYH8": GetStripeSchoolCode = "61098"        ' (SPNE)
            Case "po_1S7nNGH0skVHdzk45X5kfgEN": GetStripeSchoolCode = "35436"        ' (TUCP)
            
        ' ==============================
        '             2025.10
        ' ==============================
            Case "po_1SIfhQQhnfWS9EivT5FnnEPI": GetStripeSchoolCode = "61369"        ' (AUS)
            Case "po_1SIfcPQYaS7mQRKSmpTPeVXz": GetStripeSchoolCode = "61096"        ' (AUSP)
            Case "po_1SIfeaH9wkJRSqOaaZXTn2so": GetStripeSchoolCode = "50884"        ' (BBRM)
            Case "po_1SIfedHoHDbMF2midJy2LXld": GetStripeSchoolCode = "35412"        ' (BCSI)
            Case "po_1SIfe9H41tYMeeSKeZsXp1Vc": GetStripeSchoolCode = "61367"        ' (BRMC)
            Case "po_1SIfaRHzt9XZ7h60RxYBBaDU": GetStripeSchoolCode = "35419"        ' (CHD)
            Case "po_1SIfi7QjF0hKzNBk4eff5KcN": GetStripeSchoolCode = "35422"        ' (CHPS)
            Case "po_1SIfh1Qeaxl1oH0S0f9keHfb": GetStripeSchoolCode = "61773"        ' (CPK)
            Case "po_1SIfZhQePufX5iVQHx0xrlSD": GetStripeSchoolCode = "35425"        ' (FLG)
            Case "po_1SIfalHgYDI9wgzhBRvKCNcI": GetStripeSchoolCode = "35426"        ' (GDY)
            Case "po_1SIfeiQZOZg0F0QAiY2ADXrx": GetStripeSchoolCode = "61623"        ' (JLJ)
            Case "po_1SIfcMH46HhwWaTCIBCa9NMY": GetStripeSchoolCode = "35427"        ' (MES)
            Case "po_1SIfgYHvTxXqOAZsZKjsWcpR": GetStripeSchoolCode = "35428"        ' (OV)
            Case "po_1SIfgpHbVYyrxVVwrCRvgjzW": GetStripeSchoolCode = "35429"        ' (PEO)
            Case "po_1SIfZrH0K05F9mgHnTA7MWZg": GetStripeSchoolCode = "61621"        ' (PFL)
            Case "po_1SIfgAHePa3cGmcJ45DeiadP": GetStripeSchoolCode = "35431"        ' (PHXC)
            Case "po_1SIhGxHBO4JggQcHS3IgC59R": GetStripeSchoolCode = "62053"        ' (PHXN)
            Case "po_1SIfaIHW8Eo1hm07nvr920IK": GetStripeSchoolCode = "38346"        ' (PHXS)
            Case "po_1SIh8bQahEPZfhTMQDlqCPHR": GetStripeSchoolCode = "62375"        ' (PLN)
            Case "po_1SIhKKQkAnjuwifxRS6QTw1L": GetStripeSchoolCode = "62376"        ' (RCH)
            Case "po_1SIfhqHC1Kl6KxgMZpIzTksG": GetStripeSchoolCode = "35441"        ' (SAMC)
            Case "po_1SIfcZHZSXdjvotcA5xveTeq": GetStripeSchoolCode = "38511"        ' (SAS)
            Case "po_1SIfexQgoLYKu1DKzF4AaArg": GetStripeSchoolCode = "61098"        ' (SPNE)
            Case "po_1SIfamHWV520qxl3WxeV7gVc": GetStripeSchoolCode = "35437"        ' (TUCN)
            Case "po_1SIfd4H0skVHdzk4YKFw8ITB": GetStripeSchoolCode = "35436"        ' (TUCP)

        ' ==============================
        '             2025.11
        ' ==============================
            Case "po_1SUyzaQXwlouwMgWS0gChHaa": GetStripeSchoolCode = "35418"        ' (AHW)
            Case "po_1SUcqpQhnfWS9Eiv6OBSKEAo": GetStripeSchoolCode = "61369"        ' (AUS)
            Case "po_1SUG2bQYaS7mQRKSvLYg8nvA": GetStripeSchoolCode = "61096"        ' (AUSP)
            Case "po_1SUG2bQYaS7mQRKSvLYg8nvA": GetStripeSchoolCode = "61096"        ' (AUSP)
            Case "po_1SUGAbH9wkJRSqOaolsjvkqc": GetStripeSchoolCode = "50884"        ' (BBRM)
            Case "po_1SUFugHoHDbMF2midH3m8REt": GetStripeSchoolCode = "35412"        ' (BCSI)
            Case "po_1SUFtRQYL6jHaUXMfrIVAxTH": GetStripeSchoolCode = "61620"        ' (BEN)
            Case "po_1SUG4oH41tYMeeSK5KD4FfWo": GetStripeSchoolCode = "61367"        ' (BRMC)
            Case "po_1SUGLhHzt9XZ7h60btrhKmKB": GetStripeSchoolCode = "35419"        ' (CHD)
            Case "po_1SUFzlH837k8ldes49uEhmSB": GetStripeSchoolCode = "35468"        ' (CHPN)
            Case "po_1SUGGkQjF0hKzNBkVDrMlTTY": GetStripeSchoolCode = "35422"        ' (CHPS)
            Case "po_1SUGA0Qeaxl1oH0S0ML2X8Ra": GetStripeSchoolCode = "61773"        ' (CPK)
            Case "po_1SUcimHnTW3OArEtfh0UQmx7": GetStripeSchoolCode = "60203"        ' (DC)
            Case "po_1SUceGQePufX5iVQywUTPela": GetStripeSchoolCode = "35425"        ' (FLG)
            Case "po_1SUzTJHgYDI9wgzhEPqSmf2Q": GetStripeSchoolCode = "35426"        ' (GDY)
            Case "po_1SUzDNHUkZTU2PVIL5AXhkXM": GetStripeSchoolCode = "35440"        ' (GDYP)
            Case "po_1SUFupQZOZg0F0QA31BIcWWp": GetStripeSchoolCode = "61623"        ' (JLJ)
            Case "po_1SUys0H46HhwWaTC1R3gLR6Z": GetStripeSchoolCode = "35427"        ' (MES)
            Case "po_1SUGPXHvTxXqOAZsTXbknKq0": GetStripeSchoolCode = "35428"        ' (OV)
            Case "po_1SUytTHZjj5CWGhpqxc4dP2v": GetStripeSchoolCode = "35439"        ' (OVP)
            Case "po_1SUGKMHbVYyrxVVwelJCcwY7": GetStripeSchoolCode = "35429"        ' (PEO)
            Case "po_1SUGMIHv5cEoCFRKNpGdhW3k": GetStripeSchoolCode = "38139"        ' (PEOP)
            Case "po_1SUccEH0K05F9mgHOKg04ZsZ": GetStripeSchoolCode = "61621"        ' (PFL)
            Case "po_1SUGbUQVbVEP5TMTDE5N3qMx": GetStripeSchoolCode = "35430"        ' (PHX)
            Case "po_1SUG4GHePa3cGmcJYSj00lgK": GetStripeSchoolCode = "35431"        ' (PHXC)
            Case "po_1SUGRxHBO4JggQcHRwWdAany": GetStripeSchoolCode = "62053"        ' (PHXN)
            Case "po_1SUGKHHFL7fIaoM6UKLwgRwW": GetStripeSchoolCode = "50885"        ' (PHXP)
            Case "po_1SUGHDHW8Eo1hm07TmNAzxGp": GetStripeSchoolCode = "38346"        ' (PHXS)
            Case "po_1SUGXxQahEPZfhTMoI69HgA1": GetStripeSchoolCode = "62375"        ' (PLN)
            Case "po_1SUG7HHkv0Z5SiNeiLJyTtjH": GetStripeSchoolCode = "35432"        ' (PRE)
            Case "po_1SUGOVQkAnjuwifxRhjsWTJS": GetStripeSchoolCode = "62376"        ' (RCH)
            Case "po_1SUGCnHC1Kl6KxgMpILdH2Zj": GetStripeSchoolCode = "35441"        ' (SAMC)
            Case "po_1SUG4PHRJOrVpM8z4qqarlmP": GetStripeSchoolCode = "35442"        ' (SANC)
            Case "po_1SUFsmHQOHalnvHrow9wOgE5": GetStripeSchoolCode = "61097"        ' (SANE)
            Case "po_1SUGN7HZSXdjvotc5gwicxue": GetStripeSchoolCode = "38511"        ' (SAS)
            Case "po_1SUci0HcxeuzFJlhBz7YEFtw": GetStripeSchoolCode = "35433"        ' (SCD)
            Case "po_1SUG36HmHwPibaAkTn8jD36o": GetStripeSchoolCode = "35435"        ' (SCPE)
            Case "po_1SUz9uHbkaoiiKNoi5bgDpzr": GetStripeSchoolCode = "50886"        ' (SCPW)
            Case "po_1SUGZmQgoLYKu1DK4ZnXVIuA": GetStripeSchoolCode = "61098"        ' (SPNE)
            Case "po_1SUGWEHWV520qxl3T0J1fPOo": GetStripeSchoolCode = "35437"        ' (TUCN)
            Case "po_1SUGEKH0skVHdzk4iyTvNJXN": GetStripeSchoolCode = "35436"        ' (TUCP)
        
        ' ==============================
        '             Not Found
        ' ==============================
            Case Else
                GetStripeSchoolCode = "NONE"
    
    End Select

End Function
