CREATE TABLE role_awards (
    award_id INTEGER PRIMARY KEY AUTOINCREMENT,
    role_name TEXT NOT NULL,
    award_name TEXT NOT NULL,
    award_category TEXT NOT NULL
);

INSERT INTO role_awards (role_name, award_name, award_category) VALUES
-- Actors
('Actor', 'Academy Awards', 'Best Actor'),
('Actor', 'Academy Awards', 'Best Supporting Actor'),
('Actor', 'Golden Globe Awards', 'Best Actor'),
('Actor', 'Screen Actors Guild Awards', 'Outstanding Performance'),

-- Producers
('Producer', 'Academy Awards', 'Best Picture'),
('Producer', 'Producers Guild Awards', 'Best Theatrical Motion Picture'),
('Producer', 'Golden Globe Awards', 'Best Motion Picture'),

-- Directors
('Director', 'Academy Awards', 'Best Director'),
('Director', 'Directors Guild of America Awards', 'Outstanding Directorial Achievement'),
('Director', 'BAFTA Awards', 'Best Director'),

-- Writers
('Writer', 'Academy Awards', 'Best Original Screenplay'),
('Writer', 'Academy Awards', 'Best Adapted Screenplay'),
('Writer', 'Writers Guild Awards', 'Best Screenplay'),

-- Music
('Composer', 'Academy Awards', 'Best Original Score'),
('Composer', 'Golden Globe Awards', 'Best Original Score'),

-- Technical / Crew
('Cinematographer', 'Academy Awards', 'Best Cinematography'),
('Editor', 'Academy Awards', 'Best Film Editing'),
('Production Designer', 'Academy Awards', 'Best Production Design'),
('Costume Designer', 'Academy Awards', 'Best Costume Design'),
('Visual Effects Artist', 'Academy Awards', 'Best Visual Effects'),
('Sound Team', 'Academy Awards', 'Best Sound');
