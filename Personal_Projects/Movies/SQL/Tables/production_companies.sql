CREATE TABLE production_companies (
    company_id INTEGER PRIMARY KEY AUTOINCREMENT,
    company_name TEXT NOT NULL,
    production_start INTEGER,
    production_end INTEGER,
    movies_produced TEXT
);

INSERT INTO production_companies (company_name, production_start, production_end, movies_produced) VALUES
-- Classic / Major Studios
('Universal Pictures', 1912, NULL, '1000+'),
('Paramount Pictures', 1912, NULL, '1000+'),
('Warner Bros. Pictures', 1923, NULL, '1000+'),
('Walt Disney Pictures', 1923, NULL, '1000+'),
('Columbia Pictures (Sony)', 1918, NULL, '1000+'),
('Metro-Goldwyn-Mayer (MGM)', 1924, 1986, '500+'),
('RKO Radio Pictures', 1929, 1959, '300+'),
('20th Century Fox', 1935, 2019, '500+'),
('United Artists', 1919, 1981, '600+'),

-- Mini-majors / Mid-size (historical → modern)
('DreamWorks Pictures', 1994, NULL, '150+'),
('Lionsgate Films', 1962, NULL, '300+'),
('Miramax Films', 1979, NULL, '700+'),
('New Line Cinema', 1967, 2019, '500+'),
('Orion Pictures', 1978, 1997, '200+'),
('TriStar Pictures', 1982, NULL, '200+'),
('Summit Entertainment', 1991, 2016, '100+'),
('Focus Features', 2002, NULL, '100+'),
('A24', 2012, NULL, '150+'),
('Blumhouse Productions', 2000, NULL, '100+'),
('Relativity Media', 2004, 2018, '70+'),
('STX Entertainment', 2014, 2023, '50+'),
('Legendary Pictures', 2000, NULL, '60+'),
('Skydance Media', 2010, NULL, '50+'),
('Village Roadshow Pictures', 1997, NULL, '100+'),
('Working Title Films', 1983, NULL, '130+'),
('StudioCanal', 1988, NULL, '300+'),
('Eon Productions', 1961, NULL, '25+'),
('Amblin Entertainment', 1981, NULL, '120+');
