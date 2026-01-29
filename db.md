# DB

### 1. users
CREATE TABLE users (
  id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,

  email VARCHAR(191) NOT NULL,
  username VARCHAR(120) NOT NULL,
  password_hash VARCHAR(255) NOT NULL,

  created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,

  PRIMARY KEY (id),
  UNIQUE KEY uq_users_email (email),
  UNIQUE KEY uq_users_username (username)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;


### 2. user_mailboxes
CREATE TABLE user_mailboxes (
  id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
  user_id BIGINT UNSIGNED NOT NULL,

  mailbox_email VARCHAR(191) NOT NULL,
  tenant_id VARCHAR(64) NOT NULL DEFAULT '',
  graph_user_id VARCHAR(128) NOT NULL,

  access_token MEDIUMTEXT DEFAULT NULL,
  refresh_token MEDIUMTEXT DEFAULT NULL,
  expires_at DATETIME DEFAULT NULL,
  scopes TEXT DEFAULT NULL,

  created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,

  PRIMARY KEY (id),

  UNIQUE KEY uq_user_graph (user_id, tenant_id, graph_user_id),
  UNIQUE KEY uq_user_mailbox_email (user_id, mailbox_email),
  UNIQUE KEY uq_global_graph (tenant_id, graph_user_id),

  KEY idx_user_id (user_id),
  KEY idx_mailbox_email (mailbox_email),

  CONSTRAINT fk_user_mailboxes_user
    FOREIGN KEY (user_id) REFERENCES users(id)
    ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;



### users
Column                Type                  Null        Key       Default             Notes
id                    BIGINT UNSIGNED       NO          PRI       —                   auto_increment
email                 VARCHAR(191)          NO          UNI       —                   login email
username              VARCHAR(120)          NO          UNI       —                   display/handle
password_hash         VARCHAR(255)          NO                    —                   bcrypt/argon hash
created_at            TIMESTAMP             NO                    CURRENT_TIMESTAMP       
updated_at            TIMESTAMP             NO                    CURRENT_TIMESTAMP (on update)


### user_mailboxes
Column                Type                  Null        Key       Default             Notes
id                    BIGINT UNSIGNED       NO          PRI       —                   auto_increment
user_id               BIGINT UNSIGNED       NO          MUL       —                   FK → users.id
mailbox_email         VARCHAR(191)          NO          MUL       —                   connected mailbox UPN/mail
tenant_id             VARCHAR(64)           NO          UNI       ''                  tid claim
graph_user_id         VARCHAR(128)          NO          UNI       —                   /me id (tenant-scoped)
access_token          MEDIUMTEXT            YES                   NULL                optional (you can omit storing this)
refresh_token         MEDIUMTEXT            YES                   NULL                required for long-lived access
expires_at            DATETIME              YES                   NULL                UTC
scopes                TEXT                  YES                   NULL                stored scopes string
created_at            TIMESTAMP             NO                    CURRENT_TIMESTAMP       
updated_at            TIMESTAMP             NO                    CURRENT_TIMESTAMP (on update)
