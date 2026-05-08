using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;

namespace CMS.Data
{
    public class AppUser
    {
        public int    Id           { get; set; }
        public string Username     { get; set; } = "";
        public string PasswordHash { get; set; } = "";
        public bool   IsAdmin      { get; set; } = false;
        public bool   IsActive     { get; set; } = true;
        public DateTime CreatedAt  { get; set; } = DateTime.Now;
    }

    public static class UserRepository
    {
        private static readonly string DataFile = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "CMS", "users.json");

        private static List<AppUser> _users = new();
        private static int _nextId = 1;

        // Oturumu açan kullanıcı
        public static AppUser? CurrentUser { get; private set; }
        public static bool IsAdminLoggedIn => CurrentUser?.IsAdmin == true;

        public static void Initialize()
        {
            var dir = Path.GetDirectoryName(DataFile)!;
            if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);

            if (File.Exists(DataFile))
            {
                try
                {
                    var json = File.ReadAllText(DataFile);
                    var data = JsonSerializer.Deserialize<UserData>(json);
                    if (data != null)
                    {
                        _users  = data.Users  ?? new();
                        _nextId = data.NextId;
                    }
                }
                catch { }
            }

            // İlk çalıştırmada varsayılan admin hesabını oluştur
            if (!_users.Any())
            {
                _users.Add(new AppUser
                {
                    Id           = _nextId++,
                    Username     = "admin",
                    PasswordHash = Hash("admin123"),
                    IsAdmin      = true,
                    IsActive     = true,
                    CreatedAt    = DateTime.Now,
                });
                Save();
            }
        }

        // ── Giriş ───────────────────────────────────────────────────────
        public static bool Login(string username, string password)
        {
            var user = _users.FirstOrDefault(u =>
                u.Username.Equals(username, StringComparison.OrdinalIgnoreCase) &&
                u.PasswordHash == Hash(password) &&
                u.IsActive);

            if (user == null) return false;
            CurrentUser = user;
            return true;
        }

        public static void Logout() => CurrentUser = null;

        // ── CRUD ────────────────────────────────────────────────────────
        public static List<AppUser> GetAll() => _users.ToList();

        public static bool AddUser(string username, string password, bool isAdmin)
        {
            if (_users.Any(u => u.Username.Equals(username, StringComparison.OrdinalIgnoreCase)))
                return false;

            _users.Add(new AppUser
            {
                Id           = _nextId++,
                Username     = username,
                PasswordHash = Hash(password),
                IsAdmin      = isAdmin,
                IsActive     = true,
                CreatedAt    = DateTime.Now,
            });
            Save();
            return true;
        }

        public static void SetAdmin(int userId, bool isAdmin)
        {
            var u = _users.FirstOrDefault(x => x.Id == userId);
            if (u == null) return;
            u.IsAdmin = isAdmin;
            Save();
        }

        public static void SetActive(int userId, bool isActive)
        {
            var u = _users.FirstOrDefault(x => x.Id == userId);
            if (u == null) return;
            u.IsActive = isActive;
            Save();
        }

        public static void DeleteUser(int userId)
        {
            _users.RemoveAll(u => u.Id == userId);
            Save();
        }

        public static bool ChangePassword(int userId, string newPassword)
        {
            var u = _users.FirstOrDefault(x => x.Id == userId);
            if (u == null) return false;
            u.PasswordHash = Hash(newPassword);
            Save();
            return true;
        }

        // ── Yardımcılar ─────────────────────────────────────────────────
        private static string Hash(string input)
        {
            using var sha = System.Security.Cryptography.SHA256.Create();
            var bytes = sha.ComputeHash(System.Text.Encoding.UTF8.GetBytes(input));
            return Convert.ToHexString(bytes);
        }

        private static void Save()
        {
            var json = JsonSerializer.Serialize(new UserData { Users = _users, NextId = _nextId },
                new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(DataFile, json);
        }

        private class UserData
        {
            public List<AppUser> Users  { get; set; } = new();
            public int           NextId { get; set; } = 1;
        }
    }
}
