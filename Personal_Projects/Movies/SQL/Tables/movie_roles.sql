CREATE TABLE movie_roles (
    role_id INTEGER PRIMARY KEY AUTOINCREMENT,
    job_title TEXT NOT NULL,
    description TEXT NOT NULL
);


INSERT INTO movie_roles (job_title, description) VALUES
('Studio / Production Company', 'finances and oversees the film'),
('Producer', 'manages the project (budget, schedule, hiring)'),
('Executive Producer', 'funding / high-level oversight'),
('Director', 'creative lead; decides how the story is told'),
('Writer / Screenwriter', 'writes the story and script'),
('Animator', 'creates the character and scene animation'),
('Voice Actor', 'performs the characters'),
('Editor', 'assembles the final movie'),
('Music Composer', 'creates the score'),
('VFX / Technical Artist', 'lighting, rendering, effects');
